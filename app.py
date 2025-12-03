"""
Aplicação Flask para API GenDoc - Gestão de Demandas Redmine.
"""
from flask import Flask, jsonify, request, send_file, send_from_directory
from flask_cors import CORS
import os
import sys
import json
import tempfile
import re
from dotenv import load_dotenv
from services.redmine import buscar_demanda, formatar_dados
from services.documento import preencher_plano_trabalho

# Tenta importar redis para Vercel KV
try:
    import redis
    REDIS_AVAILABLE = True
except ImportError:
    REDIS_AVAILABLE = False

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
        
        # Carrega dados do projeto
        # Regra:
        # 1. Tenta achar um projeto cujo nomeProjeto case (case-insensitive) com o nome da demanda (project.name do Redmine)
        # 2. Se não achar, usa o primeiro projeto como fallback
        dados_projeto = {}
        projetos = carregar_projetos()
        if projetos and len(projetos) > 0:
            nome_demanda = str(dados_demanda.get('nome', '')).strip().lower()
            projeto_match = None

            if nome_demanda:
                for p in projetos:
                    nome_proj = str(p.get('nomeProjeto', '')).strip().lower()
                    # match exato ou contendo (para tolerar pequenas diferenças)
                    if nome_proj == nome_demanda or nome_demanda in nome_proj or nome_proj in nome_demanda:
                        projeto_match = p
                        break

            if projeto_match:
                dados_projeto = projeto_match
                print(f"[DEBUG] Usando projeto por nome: {dados_projeto.get('nomeProjeto', 'N/A')}")
            else:
                dados_projeto = projetos[0]
                print(f"[DEBUG] [WARN] Projeto correspondente a '{nome_demanda}' não encontrado, usando primeiro projeto: {dados_projeto.get('nomeProjeto', 'N/A')}")
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
        
        # Define o nome do arquivo de saída
        # Regra:
        # 1. Tenta usar o nome SVN do projeto (campo opcional cadastrado pelo usuário)
        # 2. Se não houver, usa o nome do projeto cadastrado
        # 3. Se ainda não houver, usa o nome vindo do Redmine
        # 4. Fallback final: "Plano_Trabalho"
        base_nome = ''
        if dados_projeto:
            base_nome = (
                str(dados_projeto.get('nomeSVN') or '').strip()
                or str(dados_projeto.get('nomeProjeto') or '').strip()
            )
        
        if not base_nome:
            base_nome = str(dados_demanda.get('nome') or '').strip()
        
        if not base_nome:
            base_nome = 'Plano_Trabalho'
        
        # Sanitiza o nome para ser um nome de arquivo seguro
        # Mantém apenas letras, números, underline, hífen e ponto; substitui o resto por underscore
        base_nome_sanitizado = re.sub(r'[^A-Za-z0-9_.-]+', '_', base_nome)
        if not base_nome_sanitizado:
            base_nome_sanitizado = 'Plano_Trabalho'
        
        download_filename = f'{base_nome_sanitizado}_{demanda_id}.docx'
        
        # Salva em arquivo temporário
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.docx')
        temp_path = temp_file.name
        temp_file.close()
        
        doc.save(temp_path)
        
        # Retorna o arquivo
        return send_file(
            temp_path,
            as_attachment=True,
            download_name=download_filename,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        ), 200
        
    except Exception as e:
        return jsonify({
            "error": "Erro ao gerar Plano de Trabalho",
            "message": str(e)
        }), 500


def get_redis_client():
    """Cria e retorna um cliente Redis para Vercel KV."""
    if not REDIS_AVAILABLE:
        return None
    
    try:
        # Vercel KV fornece essas variáveis de ambiente automaticamente
        # Para projetos antigos: KV_REST_API_URL e KV_REST_API_TOKEN
        # Para novos projetos do Marketplace: variáveis específicas do provedor
        kv_url = os.getenv('KV_REST_API_URL') or os.getenv('UPSTASH_REDIS_REST_URL')
        kv_token = os.getenv('KV_REST_API_TOKEN') or os.getenv('UPSTASH_REDIS_REST_TOKEN')
        
        if kv_url and kv_token:
            # Vercel KV/Upstash usa HTTP REST API
            return {
                'url': kv_url.rstrip('/'),
                'token': kv_token
            }
    except Exception as e:
        print(f"[WARN] Erro ao configurar Redis: {e}")
    
    return None


def get_projetos_file_path():
    """Retorna o caminho do arquivo de projetos (fallback para desenvolvimento local)."""
    return os.path.join(os.path.dirname(__file__), 'config', 'projetos.json')


def carregar_projetos():
    """Carrega a lista de projetos do Vercel KV ou arquivo JSON (fallback)."""
    # Tenta carregar do Vercel KV primeiro
    kv_client = get_redis_client()
    if kv_client:
        try:
            import requests
            # Upstash/Vercel KV REST API usa POST para todos os comandos
            # Comando: GET projetos
            response = requests.post(
                kv_client['url'],
                headers={
                    'Authorization': f"Bearer {kv_client['token']}",
                    'Content-Type': 'application/json'
                },
                json=['GET', 'projetos'],
                timeout=5
            )
            if response.status_code == 200:
                data = response.json()
                # Upstash retorna {'result': 'valor'} ou {'result': None}
                result = data.get('result')
                if result:
                    return json.loads(result)
        except Exception as e:
            print(f"[WARN] Erro ao carregar do KV, usando fallback: {e}")
            import traceback
            traceback.print_exc()
    
    # Fallback: carrega do arquivo JSON (desenvolvimento local)
    try:
        projetos_path = get_projetos_file_path()
        if os.path.exists(projetos_path):
            with open(projetos_path, 'r', encoding='utf-8') as f:
                projetos = json.load(f)
                return projetos
    except Exception as e:
        print(f"[ERROR] Erro ao carregar projetos do arquivo: {e}")
    
    return []


def salvar_projetos(projetos):
    """Salva a lista de projetos no Vercel KV ou arquivo JSON (fallback)."""
    # Tenta salvar no Vercel KV primeiro
    kv_client = get_redis_client()
    if kv_client:
        try:
            import requests
            # Upstash/Vercel KV REST API usa POST para todos os comandos
            # Comando: SET projetos "valor_json"
            projetos_json = json.dumps(projetos, ensure_ascii=False)
            response = requests.post(
                kv_client['url'],
                headers={
                    'Authorization': f"Bearer {kv_client['token']}",
                    'Content-Type': 'application/json'
                },
                json=['SET', 'projetos', projetos_json],
                timeout=5
            )
            if response.status_code == 200:
                print("[INFO] Projetos salvos no Vercel KV com sucesso")
                return True
            else:
                print(f"[WARN] Resposta inesperada do KV: {response.status_code} - {response.text}")
        except Exception as e:
            print(f"[WARN] Erro ao salvar no KV, usando fallback: {e}")
            import traceback
            traceback.print_exc()
    
    # Fallback: salva no arquivo JSON (desenvolvimento local)
    try:
        projetos_path = get_projetos_file_path()
        dir_path = os.path.dirname(projetos_path)
        if dir_path:
            os.makedirs(dir_path, exist_ok=True)
        
        with open(projetos_path, 'w', encoding='utf-8') as f:
            json.dump(projetos, f, ensure_ascii=False, indent=2)
        print("[INFO] Projetos salvos no arquivo local (fallback)")
        return True
    except Exception as e:
        print(f"[ERROR] Erro ao salvar projetos no arquivo: {e}")
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
            "nomeSVN": data.get('nomeSVN'),
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


@app.route('/api/projetos/<int:projeto_id>', methods=['PUT'])
def atualizar_projeto(projeto_id):
    """
    Rota para atualizar um projeto existente.

    Body (JSON):
        Mesma estrutura da criação de projeto. Somente os campos enviados serão atualizados.

    Returns:
        - 200: Projeto atualizado com sucesso
        - 400: Dados inválidos
        - 404: Projeto não encontrado
        - 500: Erro no servidor
    """
    try:
        data = request.get_json() or {}

        # Carrega projetos existentes
        projetos = carregar_projetos()

        # Encontra o projeto pelo ID
        projeto_existente = next((p for p in projetos if p.get("id") == projeto_id), None)
        if not projeto_existente:
            return jsonify({
                "error": "Projeto não encontrado",
                "message": f"Projeto com id {projeto_id} não foi encontrado."
            }), 404

        # Campos que podem ser atualizados
        campos_editaveis = [
            'nomeProjeto', 'nomeSVN', 'gestorNome', 'gestorEmail', 'gestorCelular',
            'gerenteNome', 'gerenteEmail', 'gerenteTelefone',
            'introducaoProjeto'
        ]

        # Atualiza apenas campos presentes no body
        for campo in campos_editaveis:
            if campo in data and data.get(campo) is not None:
                projeto_existente[campo] = data.get(campo)

        # Persiste alterações
        if salvar_projetos(projetos):
            return jsonify(projeto_existente), 200
        else:
            return jsonify({
                "error": "Erro ao salvar projeto",
                "message": "Não foi possível salvar o projeto atualizado. Verifique os logs do servidor."
            }), 500

    except Exception as e:
        return jsonify({
            "error": "Erro ao atualizar projeto",
            "message": str(e)
        }), 500


if __name__ == '__main__':
    # Configurações do servidor
    port = int(os.getenv('PORT', 5000))
    debug = os.getenv('FLASK_ENV') == 'development'
    
    app.run(host='0.0.0.0', port=port, debug=debug)

