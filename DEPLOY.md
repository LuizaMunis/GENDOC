# Guia de Deploy no Vercel

Este guia explica como fazer a API funcionar no Vercel.

## 📋 Pré-requisitos

1. Conta no Vercel (https://vercel.com)
2. Projeto conectado ao GitHub
3. Variáveis de ambiente configuradas

## 🔧 Configuração das Variáveis de Ambiente

No painel do Vercel, você precisa configurar as seguintes variáveis de ambiente:

1. Acesse seu projeto no Vercel
2. Vá em **Settings** → **Environment Variables**
3. Adicione as seguintes variáveis:

### Variáveis Obrigatórias:

- `REDMINE_API_KEY`: Sua chave de API do Redmine
- `REDMINE_BASE_URL`: URL base do Redmine (ex: `https://redmine.saude.gov.br`)

### Exemplo:

```
REDMINE_API_KEY=seu_token_aqui
REDMINE_BASE_URL=https://redmine.saude.gov.br
```

## 🚀 Deploy

### Opção 1: Deploy Automático (Recomendado)

1. Conecte seu repositório GitHub ao Vercel
2. O Vercel detectará automaticamente o projeto Python
3. Configure as variáveis de ambiente no painel
4. O deploy será feito automaticamente a cada push no GitHub

### Opção 2: Deploy Manual

```bash
# Instale o Vercel CLI
npm i -g vercel

# Faça login
vercel login

# Deploy
vercel

# Para produção
vercel --prod
```

## 📁 Estrutura de Arquivos

A estrutura do projeto está configurada assim:

```
gendoc/
├── api/
│   └── index.py          # Entry point para o Vercel
├── app.py                # Aplicação Flask principal
├── services/             # Serviços da aplicação
├── config/               # Arquivos de configuração
├── vercel.json           # Configuração do Vercel
├── requirements.txt      # Dependências Python
└── runtime.txt           # Versão do Python
```

## ✅ Verificação

Após o deploy, teste os endpoints:

- Health check: `https://seu-projeto.vercel.app/health`
- API Redmine: `https://seu-projeto.vercel.app/api/redmine/128910`
- Frontend: `https://seu-projeto.vercel.app/`

## 🐛 Troubleshooting

### Erro: "Module not found"
- Verifique se todas as dependências estão no `requirements.txt`
- Certifique-se de que o `api/index.py` está importando corretamente

### Erro: "REDMINE_API_KEY não configurada"
- Verifique se as variáveis de ambiente foram configuradas no Vercel
- Certifique-se de que estão marcadas para o ambiente correto (Production, Preview, Development)

### Erro: "Timeout"
- O Vercel tem limite de tempo para serverless functions (10s no plano gratuito)
- Para operações longas, considere usar background jobs

## 📚 Recursos

- [Documentação Vercel Python](https://vercel.com/docs/concepts/functions/serverless-functions/runtimes/python)
- [Documentação Flask](https://flask.palletsprojects.com/)

