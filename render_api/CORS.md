# CORS — Carregar GeoJSON/API em site externo (Locaweb, etc.)

Para carregar o GeoJSON (ou chamar a API) em um mapa (Leaflet, Google Maps) ou frontend hospedado **em outro domínio** (ex.: site na Locaweb), o navegador exige que a API permita essa origem (CORS). Caso contrário, a requisição é bloqueada.

## Rotas que devolvem dados (onde o CORS importa)

As rotas **`/web/*`** (ex.: `/web/geojson`) devolvem **páginas HTML** (front-end com formulários e scripts). O CORS é aplicado a todas as respostas, mas em testes com `Invoke-WebRequest` em uma URL que retorna HTML o foco costuma ser a “navegação”, não os headers de dados; o que o site externo vai chamar são as **rotas de API/dados**:

| Rota | Retorno | Autenticação |
|------|---------|--------------|
| `GET /api/config` | JSON (lotes, modalidades, etc.) | Não |
| `GET /outputs/{nome_arquivo}` | Arquivo (GeoJSON, PDF, XLSX, etc.) | Sim (cookie/Bearer) |
| `POST /gerar-relatorio-progresso` | JSON (stream de progresso) | Sim |
| `POST /geojson/upload` | JSON | Sim |

Para **testar CORS** no PowerShell, use um endpoint que retorna JSON ou arquivo, por exemplo:

```powershell
$headers = @{ "Origin" = "https://www.gestao-rodovias.com.br" }
$r = Invoke-WebRequest -Uri "https://www.gestao-rodovias.com.br/api/config" -Method Get -Headers $headers -UseBasicParsing
$r.Headers['Access-Control-Allow-Origin']
$r.Headers['Access-Control-Allow-Credentials']
```

Se o CORS estiver correto, deve aparecer a origem e `true`. Para testar download de GeoJSON (com login), use `GET /outputs/nome-do-arquivo.geojson` com cookie ou header `Authorization: Bearer <token>`.

## Configuração

Defina a variável de ambiente **`ARTESP_CORS_ORIGINS`** no ambiente onde a API está rodando (Render, Locaweb VPS, etc.).

### Lista explícita de origens (recomendado em produção)

Valor: URLs do site que vai consumir a API, separadas por vírgula, **sem espaço** após a vírgula.

Exemplos:

```env
# Um site
ARTESP_CORS_ORIGINS=https://www.seudominio.com.br

# Vários sites (com e sem www, Locaweb)
ARTESP_CORS_ORIGINS=https://www.seudominio.com.br,https://seudominio.com.br,https://meusite.locaweb.com.br
```

Use exatamente a URL que aparece no navegador (esquema e domínio).

### Aceitar qualquer origem (apenas testes)

```env
ARTESP_CORS_ORIGINS=*
```

A API reflete o header `Origin` da requisição, permitindo que qualquer domínio chame a API. Não use em produção se houver dados sensíveis.

## Backup da configuração (Render)

Valor em uso no Render para o site em produção. Use esta string exata ao criar um novo serviço ou restaurar a config:

```
ARTESP_CORS_ORIGINS=https://www.gestao-rodovias.com.br,https://gestao-rodovias.com.br
```

## Onde definir

- **Render:** Dashboard do serviço → Environment → Add Variable → `ARTESP_CORS_ORIGINS` → valor acima → Save e redeploy.
- **Locaweb VPS:** no `.env` do container ou no `docker run` / `docker-compose`: `-e ARTESP_CORS_ORIGINS=...`.

## Exemplo no front (site na Locaweb)

```javascript
// Carregar GeoJSON no Leaflet a partir da API no Render
fetch('https://sua-api.onrender.com/outputs/meu-arquivo.geojson', {
  credentials: 'include'  // se precisar enviar cookie de autenticação
})
  .then(r => r.json())
  .then(geojson => {
    L.geoJSON(geojson).addTo(map);
  });
```

O domínio do site onde esse código roda deve estar em `ARTESP_CORS_ORIGINS` na API.
