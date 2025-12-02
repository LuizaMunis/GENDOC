"""
Serviço para buscar dados de demandas no Redmine via API JSON do Redmine.
"""
import os
import requests
from typing import Dict, Optional, Any


def buscar_demanda(demanda: str) -> Optional[Dict[str, Any]]:
    """
    Busca dados de uma demanda específica no Redmine via API JSON.
    
    Args:
        demanda: Número da demanda (ID) a ser buscada
        
    Returns:
        Dicionário com o JSON bruto retornado pelo Redmine ou None se não encontrada
        
    Raises:
        Exception: Em caso de erro na requisição ou processamento
    """
    api_key = os.getenv('REDMINE_API_KEY')
    if not api_key:
        raise ValueError("REDMINE_API_KEY não configurada nas variáveis de ambiente")
    
    # Base do Redmine - pode ser configurada via variável de ambiente
    # Exemplo: https://redmine.saude.gov.br
    base_url = os.getenv('REDMINE_BASE_URL', 'https://redmine.saude.gov.br')

    # Monta a URL do endpoint JSON para a demanda específica
    # Exemplo: https://redmine.saude.gov.br/issues/<demanda>.json?include=relations,children&key=<API_KEY>
    url = f"{base_url.rstrip('/')}/issues/{demanda}.json"
    params = {
        "include": "relations,children",
        "key": api_key,
    }
    
    try:
        response = requests.get(url, params=params, timeout=30)

        if response.status_code == 404:
            # Demanda não existe no Redmine
            return None

        response.raise_for_status()

        return response.json()

    except requests.exceptions.RequestException as e:
        raise Exception(f"Erro ao acessar API do Redmine: {str(e)}")
    except Exception as e:
        raise Exception(f"Erro inesperado: {str(e)}")


def _get_custom_field(issue: Dict[str, Any], field_name: str) -> str:
    """
    Busca o valor de um custom field pelo nome dentro do JSON do Redmine.
    Tenta diferentes variações do nome (case-insensitive, com/sem espaços).
    """
    custom_fields = issue.get("custom_fields", [])
    
    # Tenta busca exata primeiro
    for campo in custom_fields:
        nome_campo = campo.get("name", "")
        if nome_campo == field_name:
            valor = campo.get("value", "") or ""
            return str(valor).strip()
    
    # Tenta busca case-insensitive
    for campo in custom_fields:
        nome_campo = campo.get("name", "")
        if nome_campo.lower() == field_name.lower():
            valor = campo.get("value", "") or ""
            return str(valor).strip()
    
    # Tenta busca parcial (contém)
    for campo in custom_fields:
        nome_campo = campo.get("name", "")
        if field_name.lower() in nome_campo.lower() or nome_campo.lower() in field_name.lower():
            valor = campo.get("value", "") or ""
            return str(valor).strip()
    
    return ""


def _get_relation(json_redmine: Dict[str, Any], relation_type: str) -> str:
    """
    Busca o ID de uma relação específica (ex: "relates", "duplicates", "blocks").
    Os dados podem estar em relations ou children.
    """
    issue = json_redmine.get("issue", {})
    
    # Busca em relations
    relations = issue.get("relations", [])
    for rel in relations:
        if rel.get("relation_type") == relation_type:
            return str(rel.get("issue_id", "")).strip()
    
    # Busca em children (filhos da demanda)
    children = issue.get("children", [])
    if children:
        # Pode haver múltiplos filhos, retorna o primeiro ou todos separados por vírgula
        ids = [str(child.get("id", "")) for child in children if child.get("id")]
        if ids:
            return ",".join(ids)
    
    return ""


def _buscar_sprint_detalhes(sprint_id: str) -> Dict[str, str]:
    """
    Busca os detalhes de uma Sprint específica no Redmine.
    Retorna um dicionário com os campos da Sprint.
    """
    api_key = os.getenv('REDMINE_API_KEY')
    base_url = os.getenv('REDMINE_BASE_URL', 'https://redmine.saude.gov.br')
    
    url = f"{base_url.rstrip('/')}/issues/{sprint_id}.json"
    params = {
        "key": api_key,
    }
    
    try:
        response = requests.get(url, params=params, timeout=30)
        if response.status_code == 404:
            return {}
        response.raise_for_status()
        sprint_data = response.json().get("issue", {})
        
        # Extrai os campos da Sprint
        valor_unitario = _get_custom_field(sprint_data, "Valor Unitário")
        valor_fase = _get_custom_field(sprint_data, "Valor da Fase")
        tipo_sprint = _get_custom_field(sprint_data, "Tipo de Sprint")
        hst = _get_custom_field(sprint_data, "Tempo Estimado (HST)")
        
        return {
            "valor_h_sprint": valor_unitario,
            "valor_total": valor_fase,
            "tipo": tipo_sprint,
            "hst": hst,
        }
    except Exception:
        return {}


def _navegar_children(children: list) -> list:
    """
    Navega recursivamente pelos children para encontrar TODAS as combinações PT-OS-Sprint.
    Retorna uma lista de dicionários, cada um com pt, os e sprint.
    """
    linhas = []
    
    for child in children:
        tracker = child.get("tracker", {})
        tracker_name = tracker.get("name", "") if isinstance(tracker, dict) else ""
        pt_id = str(child.get("id", "")).strip()
        
        # Primeiro nível: Plano de Trabalho (PT)
        if tracker_name == "Plano de Trabalho":
            # Segundo nível: Proposta de OS (dentro do PT)
            filhos_pt = child.get("children", [])
            for filho_pt in filhos_pt:
                tracker_filho = filho_pt.get("tracker", {})
                tracker_filho_name = tracker_filho.get("name", "") if isinstance(tracker_filho, dict) else ""
                
                if tracker_filho_name == "Proposta de OS":
                    os_id = str(filho_pt.get("id", "")).strip()
                    
                    # Terceiro nível: Sprint (dentro da OS) - BUSCA TODAS AS SPRINTS
                    filhos_os = filho_pt.get("children", [])
                    for filho_os in filhos_os:
                        tracker_sprint = filho_os.get("tracker", {})
                        tracker_sprint_name = tracker_sprint.get("name", "") if isinstance(tracker_sprint, dict) else ""
                        
                        if tracker_sprint_name == "Sprint":
                            sprint_id = str(filho_os.get("id", "")).strip()
                            # Adiciona uma linha para cada Sprint encontrada
                            linhas.append({
                                "pt": pt_id,
                                "os": os_id,
                                "sprint": sprint_id,
                            })
    
    return linhas


def formatar_dados(json_redmine: Dict[str, Any]) -> list:
    """
    Formata o JSON bruto do Redmine no formato esperado pela aplicação GenDoc.
    Retorna uma LISTA de linhas, uma para cada Sprint encontrada.
    
    Estrutura hierárquica esperada:
    - Demanda (raiz)
      - Plano de Trabalho (PT)
        - Proposta de OS (OS)
          - Sprint (pode haver múltiplas)
    
    Estrutura de saída (lista de objetos):
    [
      {
        "demanda": "128910",
        "pt": "129199",
        "os": "129200",
        "sprint": "129201",
        "tipo": "Manutenção",
        "nome": "Subproj GAL",
        "hst": "160",
        "valor_h_sprint": "R$ 244,67",
        "valor_total": "R$ 39.147,20",
        "valor_demanda": "R$ 78.294,40"
      },
      ...
    ]
    """
    issue = json_redmine.get("issue", {})
    
    # Dados da demanda principal (comuns a todas as linhas)
    demanda_id = str(issue.get("id", "")).strip()
    
    # Nome do projeto
    project = issue.get("project", {})
    nome = project.get("name", "") if isinstance(project, dict) else ""
    
    # Valor da Demanda (custom field)
    valor_demanda_raw = _get_custom_field(issue, "Valor da Demanda")
    
    # Formata valores monetários
    def formatar_moeda(valor: str) -> str:
        """Formata um valor numérico como moeda brasileira."""
        if not valor:
            return ""
        try:
            valor_float = float(valor)
            return f"R$ {valor_float:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except (ValueError, TypeError):
            return valor
    
    valor_demanda = formatar_moeda(valor_demanda_raw) if valor_demanda_raw else ""
    
    # Navega pelos children para encontrar TODAS as combinações PT-OS-Sprint
    children = issue.get("children", [])
    linhas_encontradas = _navegar_children(children)
    
    # Para cada linha encontrada, busca os detalhes da Sprint e cria o objeto final
    resultado = []
    for linha in linhas_encontradas:
        sprint_id = linha.get("sprint", "")
        
        # Busca os detalhes da Sprint
        sprint_detalhes = {}
        if sprint_id:
            sprint_detalhes = _buscar_sprint_detalhes(sprint_id)
        
        # Extrai os valores da Sprint
        valor_h_sprint_raw = sprint_detalhes.get("valor_h_sprint", "")
        valor_total_raw = sprint_detalhes.get("valor_total", "")
        tipo = sprint_detalhes.get("tipo", "")
        hst = sprint_detalhes.get("hst", "")
        
        valor_h_sprint = formatar_moeda(valor_h_sprint_raw) if valor_h_sprint_raw else ""
        valor_total = formatar_moeda(valor_total_raw) if valor_total_raw else ""
        
        # Cria uma linha completa para esta Sprint
        resultado.append({
            "demanda": demanda_id,
            "pt": linha.get("pt", ""),
            "os": linha.get("os", ""),
            "sprint": sprint_id,
            "tipo": tipo,
            "nome": nome,
            "hst": hst,
            "valor_h_sprint": valor_h_sprint,
            "valor_total": valor_total,
            "valor_demanda": valor_demanda,
        })
    
    return resultado

