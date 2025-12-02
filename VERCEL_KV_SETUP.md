# Como Configurar Vercel KV (Redis) para o GenDoc

## 📋 Passo a Passo

### 1. Criar o Vercel KV no Dashboard

1. Acesse: https://vercel.com/dashboard
2. Selecione seu projeto **gendoc**
3. Vá em **Storage** (ou **Integrations** → **Add Integration**)
4. Procure por **KV** ou **Redis**
5. Clique em **Create** ou **Add**

**Nota:** Se não encontrar "KV", procure por integrações Redis no Marketplace do Vercel (Upstash Redis, Redis Cloud, etc.)

### 2. Configuração Automática

O Vercel automaticamente adiciona as seguintes variáveis de ambiente:
- `KV_REST_API_URL` (ou `UPSTASH_REDIS_REST_URL`)
- `KV_REST_API_TOKEN` (ou `UPSTASH_REDIS_REST_TOKEN`)

### 3. Verificar Variáveis de Ambiente

1. No painel do Vercel, vá em **Settings** → **Environment Variables**
2. Verifique se as variáveis acima estão presentes
3. Se não estiverem, adicione manualmente (os valores estarão na documentação do KV/Redis que você criou)

### 4. Fazer Redeploy

Após configurar o KV:
1. Vá em **Deployments**
2. Clique nos três pontos (⋯) do último deploy
3. Clique em **Redeploy**
4. Ou faça um novo commit/push no GitHub

## ✅ Como Funciona

O código agora:
1. **Tenta usar Vercel KV primeiro** - Se as variáveis de ambiente estiverem configuradas
2. **Faz fallback para arquivo local** - Se KV não estiver disponível (desenvolvimento local)

## 🧪 Testar

Após configurar:
1. Acesse: `https://gendoc-livid.vercel.app`
2. Tente adicionar um novo projeto
3. Verifique se foi salvo corretamente
4. Recarregue a página e veja se o projeto persiste

## 🔍 Troubleshooting

### Erro: "Erro ao salvar projeto"
- Verifique se o KV foi criado corretamente
- Verifique se as variáveis de ambiente estão configuradas
- Veja os logs no Vercel (Deployments → Logs)

### Projetos não persistem
- Verifique se o KV está ativo no painel do Vercel
- Verifique se as variáveis de ambiente estão corretas
- Faça um redeploy após configurar o KV

### Funciona localmente mas não no Vercel
- Certifique-se de que o KV foi criado no projeto correto do Vercel
- Verifique se as variáveis estão marcadas para **Production**, **Preview** e **Development**

## 📚 Recursos

- [Documentação Vercel Storage](https://vercel.com/docs/storage)
- [Vercel Marketplace - Redis](https://vercel.com/marketplace?category=databases)

