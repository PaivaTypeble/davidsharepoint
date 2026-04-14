# DavidSharePoint

API .NET 10 em ASP.NET Core com vertical slicing para resolver um URL de SharePoint e devolver apenas os nomes de todos os ficheiros, sem fazer download.

## O que existe

- HTTP API com endpoint `POST /api/sharepoint/file-names`
- HTTP API com endpoint `POST /api/sharepoint/route-document`
- Workflow N8N genérico em `sharepoint-route-document-generic-n8n-workflow.json`
- MCP HTTP endpoint em `POST /mcp`
- OpenAPI em `/openapi/v1.json`
- Scalar em `/scalar`
- Health endpoint em `/health`

## Configuração

Preencher estas chaves em `src/DavidSharePoint.Api/appsettings.Development.json` ou em variáveis de ambiente:

```json
{
  "DocumentRouting": {
    "DestinationRootFolderUrl": "<sharepoint-folder-url>",
    "MappingWorkbookUrl": "<optional-workbook-url>",
    "MappingWorkbookFileName": "Company Acronyms and NIPC.xlsx"
  },
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
- `Files.ReadWrite.All` para executar a cópia final via `POST /api/sharepoint/route-document` com `dryRun=false`

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
DocumentRouting__DestinationRootFolderUrl=<sharepoint-folder-url>
DocumentRouting__MappingWorkbookUrl=<optional-workbook-url>
MicrosoftGraph__TenantId=<tenant-id>
MicrosoftGraph__ClientId=<client-id>
MicrosoftGraph__ClientSecret=<client-secret>
```

### Endpoints esperados em produção

- `/health`
- `/openapi/v1.json`
- `/scalar`
- `/api/sharepoint/file-names`
- `/api/sharepoint/route-document`
- `/mcp`

### Notas de operação

- O container escuta em `8080`, que é o valor a usar no `Proxy Port` do EasyPanel.
- A app já trata `X-Forwarded-Proto` e `X-Forwarded-For`, por isso funciona atrás do proxy do EasyPanel sem loops de HTTPS.
- Não há storage persistente obrigatório nesta versão, porque a API não grava ficheiros nem faz downloads.
- O fluxo de routing lê texto nativo de `pdf`, `docx`, `txt` e `csv`. OCR ainda não está ligado; quando um ficheiro precisar de OCR a API devolve `ocr_required`.
- Se quiseres restringir `Scalar` e `OpenAPI` em produção, isso pode ser feito na próxima iteração.

## Exemplo HTTP

```http
POST http://localhost:5058/api/sharepoint/file-names
Content-Type: application/json

{
  "sharePointUrl": "https://contoso.sharepoint.com/sites/Finance/Shared%20Documents/Reports"
}

POST http://localhost:5058/api/sharepoint/route-document
Content-Type: application/json

{
  "sourceFileUrl": "https://contoso.sharepoint.com/sites/Finance/Shared%20Documents/invoice.pdf",
  "dryRun": true
}
```

## Workflow N8N Genérico

Existe um workflow importável em `sharepoint-route-document-generic-n8n-workflow.json` gerado a partir de `scripts/build-generic-route-n8n-workflow.ps1`.

Notas práticas desta versão:

- Usa apenas core nodes: `Manual Trigger`, `Code`, `HTTP Request` e `Extract From File`
- Não usa nodes específicos de SharePoint, Excel ou OCR
- Faz auth no Microsoft Graph, resolve URLs SharePoint, descarrega o ficheiro fonte, lê o workbook de mapping, faz matching e cria o ficheiro final
- Para `docx`, a leitura é feita via conversão Microsoft Graph para `pdf` antes da extração de texto
- O node `Extract Workbook Rows` assume a folha `Folha1`; se o workbook mudar de nome de folha, ajustar esse node depois de importar
- OCR continua fora do workflow; imagens e PDFs sem texto nativo devolvem `ocr_required`

## Estrutura

- `src/DavidSharePoint.Api/Features/SharePoint/ListFileNames`: slice HTTP + handler + tool MCP
- `src/DavidSharePoint.Api/Features/SharePoint/RouteDocument`: preview/copy do documento com base no workbook de mapping
- `src/DavidSharePoint.Api/Infrastructure/Documents`: leitura de workbook, matching e extração nativa de texto sem OCR
- `src/DavidSharePoint.Api/Infrastructure/SharePoint`: resolução de site/drive/path e navegação no Graph
- `src/DavidSharePoint.Api/Infrastructure/Graph`: token acquisition via client credentials