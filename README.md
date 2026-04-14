# DavidSharePoint

API .NET 10 em ASP.NET Core com vertical slicing para resolver um URL de SharePoint e devolver apenas os nomes de todos os ficheiros, sem fazer download.

## O que existe

- HTTP API com endpoint `POST /api/sharepoint/file-names`
- MCP HTTP endpoint em `POST /mcp`
- OpenAPI em `/openapi/v1.json`
- Scalar em `/scalar`
- Health endpoint em `/health`

## Configuração

Preencher estas chaves em `src/DavidSharePoint.Api/appsettings.Development.json` ou em variáveis de ambiente:

```json
{
  "MicrosoftGraph": {
    "TenantId": "<tenant-id>",
    "ClientId": "<client-id>",
    "ClientSecret": "<client-secret>"
  }
}
```

Permissões esperadas no Microsoft Graph para a app registration:

- `Sites.Read.All`
- `Files.Read.All`

## Arranque local

```powershell
.\scripts\run-local.ps1
```

## Deploy no Hostinger via EasyPanel

### Pré-requisitos no servidor

- VPS Linux no Hostinger
- Portas `80` e `443` abertas
- EasyPanel instalado no servidor

Se ainda não tiveres o EasyPanel instalado, usa o one-click do Hostinger a partir da documentação do EasyPanel ou faz instalação manual num VPS limpo.

### Ficheiros de deploy já preparados

- `Dockerfile`: build e runtime para .NET 10
- `.dockerignore`: reduz o contexto de build
- `easypanel.env.example`: variáveis para copiar para o painel

### Como criar o serviço no EasyPanel

1. Fazer push deste repositório para GitHub, GitLab ou outro git provider suportado.
2. No EasyPanel, criar um novo `App Service`.
3. Em `Source`, escolher o repositório e a branch.
4. Como o repositório já tem `Dockerfile`, o EasyPanel vai construir a imagem a partir dele.
5. Em `Domains & Proxy`, definir o `Proxy Port` como `8080`.
6. Em `Environment`, copiar o conteúdo de `easypanel.env.example` e preencher os valores reais.
7. Associar o domínio pretendido e fazer deploy.

### Variáveis de ambiente no EasyPanel

Usa este formato no campo `Environment`:

```env
ASPNETCORE_ENVIRONMENT=Production
MicrosoftGraph__TenantId=<tenant-id>
MicrosoftGraph__ClientId=<client-id>
MicrosoftGraph__ClientSecret=<client-secret>
```

### Endpoints esperados em produção

- `/health`
- `/openapi/v1.json`
- `/scalar`
- `/api/sharepoint/file-names`
- `/mcp`

### Notas de operação

- O container escuta em `8080`, que é o valor a usar no `Proxy Port` do EasyPanel.
- A app já trata `X-Forwarded-Proto` e `X-Forwarded-For`, por isso funciona atrás do proxy do EasyPanel sem loops de HTTPS.
- Não há storage persistente obrigatório nesta versão, porque a API não grava ficheiros nem faz downloads.
- Se quiseres restringir `Scalar` e `OpenAPI` em produção, isso pode ser feito na próxima iteração.

## Exemplo HTTP

```http
POST http://localhost:5058/api/sharepoint/file-names
Content-Type: application/json

{
  "sharePointUrl": "https://contoso.sharepoint.com/sites/Finance/Shared%20Documents/Reports"
}
```

## Estrutura

- `src/DavidSharePoint.Api/Features/SharePoint/ListFileNames`: slice HTTP + handler + tool MCP
- `src/DavidSharePoint.Api/Infrastructure/SharePoint`: resolução de site/drive/path e navegação no Graph
- `src/DavidSharePoint.Api/Infrastructure/Graph`: token acquisition via client credentials