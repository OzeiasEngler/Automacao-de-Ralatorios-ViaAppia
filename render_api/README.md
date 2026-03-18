# API ARTESP (FastAPI — Render, uso comercial)

Servidor para processamento de GeoJSON ARTESP (validação por schema + assinatura digital), com autenticação por email/senha.

## Endpoints principais

- **GET /** — Status geral da API
- **GET /web** — Interface web com layout estilo desktop (login + processamento + log)
- **GET /health** — Health check para o Render
- **POST /auth/login** — Login por email/senha (retorna token Bearer)
- **POST /auth/logout** — Invalida token atual
- **POST /processar-relatorio** — Recebe GeoJSON, valida no schema (`conserva`/`obras`), salva em `outputs/` e tenta assinar.
  - Requer header: `Authorization: Bearer <token>`
- **GET /outputs/{arquivo}** — Download do arquivo gerado (requer token Bearer)

### Exemplo de payload em `/processar-relatorio`

```json
{
  "schema": "conserva",
  "lote": "L13",
  "assinar": true,
  "nome_arquivo": "L13_relatorio.geojson",
  "geojson": {
    "type": "FeatureCollection",
    "features": []
  }
}
```

## Variáveis de ambiente de assinatura

- `ARTESP_PFX` (caminho no disco) **ou** `ARTESP_PFX_CONTENT` (Base64 do .pfx)
- `ARTESP_PFX_PASSWORD` (senha do certificado)

## Variáveis de ambiente de autenticação web

- `ARTESP_WEB_USERS` (recomendado), em um destes formatos:
  - JSON objeto: `{"usuario@empresa.com":"<senha_ou_hash>"}`
  - JSON lista: `[{"email":"usuario@empresa.com","senha":"<senha_ou_hash>"}]`
  - CSV: `usuario@empresa.com:<senha_ou_hash>,admin@empresa.com:<senha_ou_hash>`
  - formato recomendado de hash: `pbkdf2_sha256$iter$salt_b64$digest_b64`
- ou fallback:
  - `ARTESP_WEB_ADMIN_EMAIL`
  - `ARTESP_WEB_ADMIN_PASSWORD` (aceita senha pura ou hash PBKDF2)
- opcional:
  - `ARTESP_WEB_TOKEN_TTL_SECONDS` (default: `28800`)
  - `ARTESP_WEB_PBKDF2_ITERATIONS` (default: `390000`)

## Painel Admin (Dashboard de métricas)

O endpoint `GET /admin/stats` requer permissão de administrador. Configure admins via:

- **ARTESP_ADMIN_EMAILS** — lista de emails separados por vírgula, ex: `admin@empresa.com,gestor@artesp.sp.gov.br`
- **ARTESP_WEB_ADMIN_EMAIL** — fallback: um único email admin (também usado para login quando `ARTESP_WEB_USERS` não está definido)

Exemplo: `ARTESP_ADMIN_EMAILS=admin@artesp.sp.gov.br,supervisor@concessionaria.com`

## Auto-limpeza (evitar encher o disco no Render)

Arquivos em `outputs/` com mais de 24h são removidos automaticamente:

- **Startup:** limpeza na subida do servidor
- **Background:** limpeza a cada 6h (opcional)

Variáveis de ambiente:

- `ARTESP_LIMPEZA_HORAS` — idade em horas para considerar arquivo antigo (default: `24`)
- `ARTESP_LIMPEZA_INTERVALO_SEG` — intervalo em segundos para limpeza em background (default: `21600` = 6h). Use `0` para desativar.

### Gerar hash de senha (recomendado)

- `python gerar_hash_senha.py "MinhaSenhaForte123" --email "usuario@empresa.com"`
- saída inclui hash e exemplo pronto para `ARTESP_WEB_USERS`.

## Alterações no código — quando passam a valer

**O servidor carrega o código na subida.** Alterações em `nc_artesp/` ou em `render_api/` **só têm efeito depois de reiniciar o processo** (parar e subir de novo o uvicorn, ou novo deploy no Render).

- **Desenvolvimento local** (na raiz do projeto):
  ```bash
  python -m uvicorn render_api.app:app --host 127.0.0.1 --port 8000
  ```
  Ou execute `run_local.bat` (Windows). Com `--reload` para recarregar ao salvar:
  ```bash
  python -m uvicorn render_api.app:app --host 127.0.0.1 --port 8000 --reload
  ```
- No **Render:** faça um novo deploy após mudar o código.
- Ao usar M02 (Gerar Modelo Foto) ou outro módulo NC, o log da API mostra qual arquivo `.py` foi carregado; confira se é o que você editou.

## Padrão GeoJSON (layout e log)

O **GeoJSON** (`/web` — `index.html`) é a referência em produção no Render. As páginas **Fotos** e **NC** devem seguir o mesmo padrão de layout e log (ver `render_api/web/PADRAO_GEOJSON_REFERENCIA.md`). Ao alterar funções de log (`logMsg`, `limparLog`) ou o HTML do log vazio, é necessário **verificar todo o bloco**: buscar por todas as chamadas a essas funções no arquivo e em painéis relacionados para manter o comportamento consistente.

## Deploy no Render

1. [render.com](https://render.com) → New → Web Service.
2. Conecte o repositório ou envie a pasta `render_api`.
3. **Build:** `pip install -r requirements.txt`
4. **Start:** `uvicorn app:app --host 0.0.0.0 --port $PORT`
5. **Environment:** defina autenticação (`ARTESP_WEB_USERS`) e assinatura (`ARTESP_PFX`/`ARTESP_PFX_CONTENT` + `ARTESP_PFX_PASSWORD`).
6. Deploy. URL exemplo: `https://api-licenca-artesp.onrender.com`
7. Acesse a UI em `https://api-licenca-artesp.onrender.com/web`.

## Exemplo de autenticação e uso

1. Login:
   - `POST /auth/login`
   - body:
     ```json
     { "email": "usuario@empresa.com", "senha": "senhaForte123" }
     ```
2. Use o `access_token` retornado no header:
   - `Authorization: Bearer <token>`
3. Chame `POST /processar-relatorio`.

## Testes

Na **raiz** do repositório (usa `pytest.ini`):

```bash
pip install -r requirements.txt
python -m pytest
```

Ou `test_local.bat` (Windows). Alternativa: `pytest render_api/tests/ -v` a partir da pasta `render_api`.

Os testes requerem dependências completas (PyJWT, pandas, etc.). O `conftest.py` define `teste@artesp.local` / `teste123`.

**Analisar PDF localmente** (script + modo alertas com PDF antigo): ver `TESTE_ANALISAR_PDF_LOCAL.md` na raiz (`--teste-local` ou `ARTESP_NC_TESTE_LOCAL=1`).

## Excel — não apagar formatação

Templates `.xlsx` (relatório, Kcor-Kria): usar **openpyxl** com `load_workbook` e alterar células; evitar `pandas.to_excel()` por cima do mesmo ficheiro. Ver **[docs/EXCEL_PRESERVAR_FORMATACAO.md](../docs/EXCEL_PRESERVAR_FORMATACAO.md)** (modo edição, `ExcelWriter` + `overlay`).

## Pipeline NC (Não Conformidades)

A interface **NC** (`/web/nc`) suporta **etapa isolada** (uma ação por vez) e **pipeline completo** (várias etapas reutilizando o mesmo `job_id`). Contrato de integração frontend:

- **Contrato:** [CONTRATO_FRONTEND_NC.md](CONTRATO_FRONTEND_NC.md) — onde guardar `job_id` (localStorage `artesp_nc_job_id`), quando enviar `finalize=1`, como lidar com respostas em stream (`?format=zip|xlsx`) e header `X-NC-Job-Id`.
- **Download por job:** `GET /outputs/nc/{job_id}/{subpath}` (ex.: `final/nc_separados.zip`). Faz touch no job para não expirar durante o uso.
- **Status com touch:** `GET /nc/jobs/{job_id}` retorna o job e atualiza `last_access`.
