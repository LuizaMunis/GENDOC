"""
Entry point para a aplicação Flask no Vercel.
Este arquivo permite que o Flask funcione como serverless function no Vercel.
"""
import sys
import os

# Adiciona o diretório raiz ao path para importar os módulos
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

# Importa o app Flask
from app import app

# O Vercel automaticamente detecta e usa o objeto 'app' Flask
# Não precisa de handler customizado, apenas exportar o app

