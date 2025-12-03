"""
Script auxiliar para rodar a aplicação localmente.
Use: python run_local.py
"""
import os
import sys

# Verifica se o arquivo .env existe
if not os.path.exists('.env'):
    print("=" * 60)
    print("AVISO: Arquivo .env não encontrado!")
    print("=" * 60)
    print("\nPor favor, crie um arquivo .env na raiz do projeto com:")
    print("\n  REDMINE_API_KEY=sua_chave_api_redmine")
    print("  REDMINE_BASE_URL=https://redmine.saude.gov.br")
    print("  PORT=5000")
    print("  FLASK_ENV=development")
    print("\nVeja o README.md para mais informações.")
    print("=" * 60)
    
    resposta = input("\nDeseja continuar mesmo assim? (s/N): ").strip().lower()
    if resposta != 's':
        print("Encerrando...")
        sys.exit(1)

# Verifica se as dependências estão instaladas
try:
    import flask
    import flask_cors
    import dotenv
    import requests
    import docx
except ImportError as e:
    print("=" * 60)
    print("ERRO: Dependências não encontradas!")
    print("=" * 60)
    print(f"\nMódulo faltando: {e.name}")
    print("\nPor favor, execute:")
    print("  pip install -r requirements.txt")
    print("=" * 60)
    sys.exit(1)

# Importa e roda a aplicação
from app import app

if __name__ == '__main__':
    port = int(os.getenv('PORT', 5000))
    debug = os.getenv('FLASK_ENV') == 'development'
    
    print("\n" + "=" * 60)
    print("🚀 GenDoc - Gerador de Documentos")
    print("=" * 60)
    print(f"📍 Servidor rodando em: http://localhost:{port}")
    print(f"🔧 Modo debug: {'Ativado' if debug else 'Desativado'}")
    print("=" * 60)
    print("\nPressione CTRL+C para parar o servidor\n")
    
    try:
        app.run(host='127.0.0.1', port=port, debug=debug)
    except KeyboardInterrupt:
        print("\n\nServidor encerrado pelo usuário.")


