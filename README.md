# GenDoc - Gerador de Documentos de Plano de Trabalho

Aplicação Flask para busca de demandas no Redmine e geração automática de Planos de Trabalho em formato Word (.docx).

## 📋 Funcionalidades

- 🔍 Busca de demandas no Redmine via API
- 📊 Visualização de sprints e informações relacionadas
- 👥 Gerenciamento de profissionais por sprint
- 📄 Geração automática de Planos de Trabalho em Word
- 🗂️ Gerenciamento de projetos (gestores, gerentes, etc.)

## 🚀 Como Rodar Localmente

### Pré-requisitos

- Python 3.11 (verifique com `python --version`)
- pip (gerenciador de pacotes Python)

### Passo 1: Instalar Dependências

```bash
pip install -r requirements.txt
```

### Passo 2: Configurar Variáveis de Ambiente

Crie um arquivo `.env` na raiz do projeto com as seguintes variáveis:

```env
# Chave da API do Redmine (obrigatório)
REDMINE_API_KEY=sua_chave_api_redmine_aqui

# URL base do Redmine (opcional - padrão: https://redmine.saude.gov.br)
REDMINE_BASE_URL=https://redmine.saude.gov.br

# Porta do servidor Flask (opcional - padrão: 5000)
PORT=5000

# Ambiente Flask (opcional - 'development' para debug ativado)
FLASK_ENV=development
```

**Importante:** 
- Substitua `sua_chave_api_redmine_aqui` pela sua chave real da API do Redmine
- O arquivo `.env` já está no `.gitignore` e não será versionado

### Passo 3: Criar Arquivo de Configuração de Projetos (opcional)

O arquivo `config/projetos.json` será criado automaticamente quando você adicionar um projeto pela interface. Se quiser criar manualmente, crie um arquivo vazio:

```json
[]
```

### Passo 4: Executar a Aplicação

```bash
python app.py
```

A aplicação estará disponível em: **http://localhost:5000**

### Passo 5: Acessar no Navegador

Abra seu navegador e acesse: `http://localhost:5000`

## 📁 Estrutura do Projeto

```
GENDOC/
├── api/                    # Endpoint para Vercel (serverless)
│   └── index.py
├── config/                 # Arquivos de configuração
│   ├── sprints_config.json
│   └── projetos.json       # Criado automaticamente
├── services/               # Serviços da aplicação
│   ├── documento.py        # Geração de documentos Word
│   └── redmine.py          # Integração com API do Redmine
├── app.py                  # Aplicação Flask principal
├── index.html              # Interface web
├── requirements.txt        # Dependências Python
└── Modelo PT-CURSOR.docx   # Modelos Word para geração
```

## 🔧 Configurações

### Variáveis de Ambiente

| Variável | Obrigatório | Descrição | Padrão |
|----------|------------|-----------|--------|
| `REDMINE_API_KEY` | ✅ Sim | Chave da API do Redmine | - |
| `REDMINE_BASE_URL` | ❌ Não | URL base do Redmine | `https://redmine.saude.gov.br` |
| `PORT` | ❌ Não | Porta do servidor Flask | `5000` |
| `FLASK_ENV` | ❌ Não | Ambiente Flask (`development` ou `production`) | - |

### Configuração de Sprints

O arquivo `config/sprints_config.json` contém as configurações de tipos de sprint e suas atividades/entregáveis correspondentes. Você pode editá-lo conforme necessário.

## 📝 Endpoints da API

### GET `/`
Página principal (HTML)

### GET `/health`
Health check da API

### GET `/api/redmine/<demanda>`
Busca dados de uma demanda no Redmine

### POST `/api/gerar-plano-trabalho`
Gera o Plano de Trabalho em formato Word

### GET `/api/projetos`
Lista todos os projetos cadastrados

### POST `/api/projetos`
Adiciona um novo projeto

## 🛠️ Desenvolvimento

### Modo Debug

Para ativar o modo debug (recarregamento automático ao salvar arquivos), defina:

```env
FLASK_ENV=development
```

### Estrutura de Dados

#### Dados da Demanda
```json
{
  "demanda": "128910",
  "pt": "129199",
  "os": "129200",
  "sprint": "129201",
  "tipo": "Manutenção",
  "nome": "Nome do Projeto",
  "hst": "160",
  "valor_h_sprint": "R$ 244,67",
  "valor_total": "R$ 39.147,20",
  "valor_demanda": "R$ 78.294,40"
}
```

#### Dados de Profissionais
```json
{
  "sprint_id": [
    {
      "tipo": "Desenvolvedor",
      "quantidade": 1,
      "horas": 40
    }
  ]
}
```

## 🐛 Troubleshooting

### Erro: "REDMINE_API_KEY não configurada"
- Verifique se o arquivo `.env` existe na raiz do projeto
- Confirme que a variável `REDMINE_API_KEY` está definida no arquivo

### Erro: "Module not found"
- Execute `pip install -r requirements.txt` novamente
- Verifique se está usando Python 3.11

### Erro ao gerar documento Word
- Verifique se os arquivos de modelo `.docx` existem na raiz do projeto
- Confirme que os dados das sprints estão preenchidos corretamente

### Porta já em uso
- Mude a porta no arquivo `.env`: `PORT=5001`
- Ou pare o processo que está usando a porta 5000

## 📦 Dependências

- Flask 3.0.0 - Framework web
- flask-cors 4.0.0 - CORS para requisições
- python-dotenv 1.0.0 - Gerenciamento de variáveis de ambiente
- requests 2.31.0 - Requisições HTTP
- python-docx 1.1.0 - Manipulação de documentos Word
- redis 5.0.1 - Cliente Redis (usado apenas para Vercel KV)

## 📄 Licença

Este projeto é de uso interno.

## 🤝 Suporte

Para dúvidas ou problemas, verifique:
1. Os logs do servidor Flask no terminal
2. O console do navegador (F12) para erros JavaScript
3. O arquivo de configuração `.env`

