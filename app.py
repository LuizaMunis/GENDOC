"""
Aplicação Flask para API GenDoc - Gestão de Demandas Redmine.
"""
from flask import Flask, jsonify, request, send_file, send_from_directory
from flask_cors import CORS
import os
import sys
import json
import tempfile
from dotenv import load_dotenv
from services.redmine import buscar_demanda, formatar_dados
from services.documento import preencher_plano_trabalho

# Carrega variáveis de ambiente do arquivo .env
load_dotenv()

# Inicializa a aplicação Flask
app = Flask(__name__)

# Habilita CORS para permitir requisições do frontend
CORS(app)


@app.route('/')
def index():
    """Rota para servir a página HTML principal."""
    try:
        # Tenta encontrar o arquivo index.html no diretório raiz
        index_path = os.path.join(os.path.dirname(__file__), 'index.html')
        if os.path.exists(index_path):
            return send_file(index_path)
        else:
            # Se não encontrar, tenta no diretório atual
            return send_from_directory('.', 'index.html')
    except Exception as e:
        # Fallback caso não encontre o arquivo
        return f"""
        <!DOCTYPE html>
        <html>
            <head>
                <title>GenDoc API</title>
                <meta charset="utf-8">
                <style>
                    body {{ font-family: Arial, sans-serif; padding: 20px; }}
                    .status {{ color: green; font-weight: bold; }}
                    a {{ color: #0070f3; text-decoration: none; }}
                    a:hover {{ text-decoration: underline; }}
                </style>
            </head>
            <body>
                <h1>GenDoc API está funcionando! ✅</h1>
                <p class="status">Status: API Online</p>
                <p>O arquivo index.html não foi encontrado, mas a API está funcionando.</p>
                <p><strong>Erro:</strong> {str(e)}</p>
                <hr>
                <h2>Endpoints disponíveis:</h2>
                <ul>
                    <li><a href="/health">/health</a> - Health check da API</li>
                    <li><a href="/api/redmine/128910">/api/redmine/&lt;id&gt;</a> - Buscar demanda do Redmine</li>
                    <li><a href="/api/projetos">/api/projetos</a> - Listar projetos</li>
                </ul>
            </body>
        </html>
        """, 200


@app.route('/api/redmine/<demanda>', methods=['GET'])
def buscar_demanda_route(demanda):
    """
    Rota para buscar dados de uma demanda no Redmine.
    
    Args:
        demanda: Número da demanda (ID)
        
    Returns:
        JSON com os dados da demanda:
        - 200: Demanda encontrada
        - 404: Demanda não encontrada
        - 500: Erro no servidor
    """
    try:
        # Busca os dados brutos da demanda no Redmine via API JSON
        json_redmine = buscar_demanda(demanda)

        if json_redmine is None:
            return jsonify({
                "error": "Demanda não encontrada",
                "demanda": demanda
            }), 404

        # Formata os dados no padrão esperado pelo frontend
        dados_formatados = formatar_dados(json_redmine)

        return jsonify(dados_formatados), 200

    except ValueError as e:
        # Erro de configuração (ex: API key não configurada)
        return jsonify({
            "error": "Erro de configuração",
            "message": str(e)
        }), 500
        
    except Exception as e:
        # Erro genérico
        return jsonify({
            "error": "Erro ao buscar demanda",
            "message": str(e)
        }), 500


@app.route('/health', methods=['GET'])
def health_check():
    """
    Rota de health check para verificar se a API está funcionando.
    """
    # Verifica se as variáveis de ambiente estão configuradas
    redmine_key = os.getenv('REDMINE_API_KEY')
    redmine_url = os.getenv('REDMINE_BASE_URL', 'https://redmine.saude.gov.br')
    
    return jsonify({
        "status": "ok",
        "service": "GenDoc API",
        "environment": {
            "redmine_api_key_configured": bool(redmine_key),
            "redmine_base_url": redmine_url,
            "python_version": sys.version.split()[0]
        }
    }), 200


@app.route('/api/redmine/<demanda>/debug', methods=['GET'])
def debug_demanda(demanda):
    """
    Rota de debug para retornar o JSON completo do Redmine.
    Útil para identificar onde estão os dados corretos.
    """
    try:
        from services.redmine import buscar_demanda
        json_redmine = buscar_demanda(demanda)
        
        if json_redmine is None:
            return jsonify({
                "error": "Demanda não encontrada",
                "demanda": demanda
            }), 404
        
        return jsonify(json_redmine), 200
        
    except Exception as e:
        return jsonify({
            "error": "Erro ao buscar demanda",
            "message": str(e)
        }), 500


@app.route('/api/gerar-plano-trabalho', methods=['POST'])
def gerar_plano_trabalho():
    """
    Rota para gerar o Plano de Trabalho em formato Word.
    
    Recebe:
    {
        "demanda": "128910",
        "dados_demanda": {...},
        "dados_sprints": [...],
        "dados_profissionais": {...}
    }
    
    Retorna o arquivo .docx gerado.
    """
    try:
        data = request.get_json()
        
        if not data:
            return jsonify({
                "error": "Dados não fornecidos"
            }), 400
        
        demanda_id = data.get('demanda')
        dados_demanda = data.get('dados_demanda', {})
        dados_sprints = data.get('dados_sprints', [])
        dados_profissionais = data.get('dados_profissionais', {})
        
        # Log para debug
        print(f"[DEBUG] Dados recebidos:")
        print(f"  Demanda ID: {demanda_id}")
        print(f"  Dados Demanda: {dados_demanda}")
        print(f"  Dados Sprints: {dados_sprints}")
        print(f"  Dados Profissionais: {dados_profissionais}")
        
        if not demanda_id:
            return jsonify({
                "error": "ID da demanda não fornecido"
            }), 400
        
        # Determina qual modelo usar baseado no tipo da sprint
        # Se algum sprint for do tipo "Desenvolvimento", usa ModeloPT-LEO-CURSOR.docx
        # Caso contrário, usa Modelo PT-CURSOR.docx
        usar_modelo_leo = False
        for sprint in dados_sprints:
            tipo_sprint = str(sprint.get('tipo', '')).strip().lower()
            if tipo_sprint == 'desenvolvimento':
                usar_modelo_leo = True
                break
        
        if usar_modelo_leo:
            modelo_nome = 'ModeloPT-LEO-CURSOR.docx'
            print(f"[DEBUG] Usando modelo LEO (Desenvolvimento): {modelo_nome}")
        else:
            modelo_nome = 'Modelo PT-CURSOR.docx'
            print(f"[DEBUG] Usando modelo padrão: {modelo_nome}")
        
        # Caminho do modelo
        modelo_path = os.path.join(os.path.dirname(__file__), modelo_nome)
        
        if not os.path.exists(modelo_path):
            return jsonify({
                "error": f"Modelo de Plano de Trabalho não encontrado: {modelo_nome}"
            }), 404
        
        # Carrega dados do projeto (usa o primeiro projeto cadastrado por padrão)
        dados_projeto = {}
        projetos = carregar_projetos()
        if projetos and len(projetos) > 0:
            # Usa o primeiro projeto por padrão
            dados_projeto = projetos[0]
            print(f"[DEBUG] Usando projeto: {dados_projeto.get('nomeProjeto', 'N/A')}")
        else:
            print(f"[DEBUG] [WARN] Nenhum projeto encontrado no arquivo de configuração")
        
        # Gera o documento
        doc = preencher_plano_trabalho(
            modelo_path=modelo_path,
            dados_demanda=dados_demanda,
            dados_sprints=dados_sprints,
            dados_profissionais=dados_profissionais,
            dados_projeto=dados_projeto
        )
        
        # Salva em arquivo temporário
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.docx')
        temp_path = temp_file.name
        temp_file.close()
        
        doc.save(temp_path)
        
        # Retorna o arquivo
        return send_file(
            temp_path,
            as_attachment=True,
            download_name=f'Plano_Trabalho_{demanda_id}.docx',
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        ), 200
        
    except Exception as e:
        return jsonify({
            "error": "Erro ao gerar Plano de Trabalho",
            "message": str(e)
        }), 500


def get_projetos_file_path():
    """Retorna o caminho do arquivo de projetos."""
    # No Vercel, usa /tmp que é o único diretório gravável
    # Em desenvolvimento local, usa o diretório config
    if os.path.exists('/tmp'):
        return '/tmp/projetos.json'
    else:
        return os.path.join(os.path.dirname(__file__), 'config', 'projetos.json')


def carregar_projetos():
    """Carrega a lista de projetos do arquivo JSON."""
    try:
        projetos_path = get_projetos_file_path()
        
        # Se o arquivo não existe em /tmp, tenta carregar do config original
        if not os.path.exists(projetos_path):
            # Tenta carregar do arquivo original em config/ (read-only no Vercel)
            config_original = os.path.join(os.path.dirname(__file__), 'config', 'projetos.json')
            if os.path.exists(config_original):
                with open(config_original, 'r', encoding='utf-8') as f:
                    projetos = json.load(f)
                    # Se estamos em /tmp, copia para lá para poder modificar depois
                    if projetos_path.startswith('/tmp'):
                        salvar_projetos(projetos)
                    return projetos
            return []
        
        with open(projetos_path, 'r', encoding='utf-8') as f:
            projetos = json.load(f)
            return projetos
    except Exception as e:
        print(f"[ERROR] Erro ao carregar projetos: {e}")
        import traceback
        traceback.print_exc()
        return []


def salvar_projetos(projetos):
    """Salva a lista de projetos no arquivo JSON."""
    try:
        projetos_path = get_projetos_file_path()
        
        # Garante que o diretório existe (só necessário se não for /tmp)
        dir_path = os.path.dirname(projetos_path)
        if dir_path and not projetos_path.startswith('/tmp'):
            os.makedirs(dir_path, exist_ok=True)
        
        with open(projetos_path, 'w', encoding='utf-8') as f:
            json.dump(projetos, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        print(f"[ERROR] Erro ao salvar projetos: {e}")
        import traceback
        traceback.print_exc()
        return False


@app.route('/api/projetos', methods=['GET'])
def listar_projetos():
    """
    Rota para listar todos os projetos cadastrados.
    
    Returns:
        JSON com a lista de projetos:
        - 200: Lista de projetos (pode ser vazia)
        - 500: Erro no servidor
    """
    try:
        projetos = carregar_projetos()
        return jsonify(projetos), 200
    except Exception as e:
        return jsonify({
            "error": "Erro ao listar projetos",
            "message": str(e)
        }), 500


@app.route('/api/projetos', methods=['POST'])
def adicionar_projeto():
    """
    Rota para adicionar um novo projeto.
    
    Body (JSON):
        {
            "nomeProjeto": "Nome do Projeto",
            "gestorNome": "Nome do Gestor",
            "gestorEmail": "email@exemplo.com",
            "gestorCelular": "(00) 00000-0000",
            "gerenteNome": "Nome do Gerente",
            "gerenteEmail": "email@exemplo.com",
            "gerenteTelefone": "(00) 0000-0000",
            "introducaoProjeto": "Descrição do projeto"
        }
    
    Returns:
        JSON com o projeto criado:
        - 201: Projeto criado com sucesso
        - 400: Dados inválidos
        - 500: Erro no servidor
    """
    try:
        data = request.get_json()
        
        # Validação dos campos obrigatórios
        campos_obrigatorios = [
            'nomeProjeto', 'gestorNome', 'gestorEmail', 'gestorCelular',
            'gerenteNome', 'gerenteEmail', 'gerenteTelefone',
            'introducaoProjeto'
        ]
        
        for campo in campos_obrigatorios:
            if not data.get(campo):
                return jsonify({
                    "error": f"Campo obrigatório faltando: {campo}"
                }), 400
        
        # Carrega projetos existentes
        projetos = carregar_projetos()
        
        # Cria novo projeto
        novo_projeto = {
            "id": len(projetos) + 1,  # ID simples baseado no tamanho da lista
            "nomeProjeto": data.get('nomeProjeto'),
            "gestorNome": data.get('gestorNome'),
            "gestorEmail": data.get('gestorEmail'),
            "gestorCelular": data.get('gestorCelular'),
            "gerenteNome": data.get('gerenteNome'),
            "gerenteEmail": data.get('gerenteEmail'),
            "gerenteTelefone": data.get('gerenteTelefone'),
            "introducaoProjeto": data.get('introducaoProjeto')
        }
        
        # Adiciona à lista
        projetos.append(novo_projeto)
        
        # Salva no arquivo
        if salvar_projetos(projetos):
            return jsonify(novo_projeto), 201
        else:
            return jsonify({
                "error": "Erro ao salvar projeto",
                "message": "Não foi possível salvar o projeto. Verifique os logs do servidor."
            }), 500
            
    except Exception as e:
        return jsonify({
            "error": "Erro ao adicionar projeto",
            "message": str(e)
        }), 500


if __name__ == '__main__':
    # Configurações do servidor
    port = int(os.getenv('PORT', 5000))
    debug = os.getenv('FLASK_ENV') == 'development'
    
    app.run(host='0.0.0.0', port=port, debug=debug)

