# Contrato Frontend — Pipeline NC ARTESP

Documento que define como o frontend deve armazenar `job_id`, quando enviar `finalize=1` e como lidar com respostas em stream (ZIP/XLSX) mantendo o `job_id` para encadear etapas.

---

## 1. Onde armazenar `job_id`

| Onde | Uso |
|------|-----|
| **localStorage** (chave `artesp_nc_job_id`) | Persistir entre recarregamentos e abas. Ideal para o usuário poder fechar a página e continuar depois, ou para encadear etapas em chamadas separadas. |
| **Memória** (variável JS) | Sessão única, sem recarregar. Suficiente se o fluxo for linear na mesma página. |

**Recomendação:** usar **localStorage** como fonte da verdade. Ao receber um `job_id` (do JSON ou do header `X-NC-Job-Id`), gravar em `localStorage.setItem('artesp_nc_job_id', job_id)`. Ao iniciar uma **nova** pipeline (primeira etapa sem reutilizar job), chamar `localStorage.removeItem('artesp_nc_job_id')` antes do primeiro POST.

```js
const NC_JOB_KEY = 'artesp_nc_job_id';

function getNcJobId() { return localStorage.getItem(NC_JOB_KEY) || ''; }
function setNcJobId(id) { if (id) localStorage.setItem(NC_JOB_KEY, id); else localStorage.removeItem(NC_JOB_KEY); }
function clearNcJobId() { localStorage.removeItem(NC_JOB_KEY); }
```

---

## 2. Quando enviar `finalize=1`

- **Etapa isolada** (uma única chamada, ex.: só “Separar”): não é obrigatório enviar `finalize`; o backend já trata como final e retorna `finished` com retenção 24h.
- **Pipeline passo a passo:** enviar **`finalize=1`** (ou `finalize=true` conforme o backend) **apenas na última etapa** que entrega o resultado final (ex.: **Inserir número** quando for o último passo, ou **Juntar** se o usuário parar aí).

Regra prática: na **última** requisição do fluxo que o usuário está fazendo, incluir no `FormData`:

```js
formData.append('finalize', '1');   // ou true, conforme aceite do backend
```

Exemplos:

- Fluxo “Separar → Gerar modelo → Conservação → Juntar → Inserir número”: enviar `finalize=1` no **Inserir número**.
- Fluxo “Separar → Gerar modelo → Conservação → Juntar” (parar no Juntar): enviar `finalize=1` no **Juntar**.

---

## 3. Resposta padrão (JSON)

Quando o backend retorna **JSON** (sem `?format=zip` nem `?format=xlsx`), o corpo segue o formato:

```json
{
  "ok": true,
  "job_id": "nc_20260224_123456_abc123",
  "stage": "stage1",
  "artifacts": {
    "stage1": ["eaf_0/planilha.xlsx"],
    "stage2": [],
    "final": []
  },
  "download_urls": ["/outputs/nc/nc_20260224_.../final/nc_separados.zip"],
  "final_files": ["nc_separados.zip"]
}
```

**O que o front deve fazer:**

1. Guardar `job_id`: `setNcJobId(response.job_id)`.
2. Nas próximas requisições do mesmo fluxo, incluir no `FormData`: `formData.append('job_id', getNcJobId())`.
3. Para download: usar os itens de `download_urls` (path relativo ao host). Ex.: `GET ${window.location.origin}${url}` ou abrir em nova aba. O download já faz touch no job (não expira no meio do uso).

---

## 4. Modo stream (`?format=zip` ou `?format=xlsx`)

Quando o front pede arquivo em stream (para “baixar direto”):

- **URL:** adicionar query string `?format=zip` ou `?format=xlsx` ao endpoint (ex.: `POST /nc/separar?format=zip`).
- **Resposta:** o corpo é o arquivo (ZIP ou XLSX); o backend envia o **header `X-NC-Job-Id`** com o `job_id` daquele job.

**O que o front deve fazer:**

1. Após o `fetch`, ler o header: `const jobId = response.headers.get('X-NC-Job-Id')`.
2. Se existir: `setNcJobId(jobId)` para poder encadear a próxima etapa mesmo quando a resposta foi stream.
3. Tratar o corpo como binário (blob/arrayBuffer) e disparar o download como hoje.

Assim, mesmo no modo stream, o front sempre obtém o `job_id` e pode usá-lo na próxima chamada (pipeline).

---

## 5. Resumo do fluxo no front

| Momento | Ação |
|--------|------|
| Usuário inicia **novo** fluxo (primeira etapa) | `clearNcJobId()` antes do primeiro POST (opcional; sem job_id o backend cria um novo). |
| Primeira etapa (ex.: Separar) | POST **sem** `job_id`. Backend retorna JSON (ou stream + `X-NC-Job-Id`). Guardar `job_id`. |
| Etapas seguintes (ex.: Gerar modelo, Conservação, Juntar, Inserir número) | POST **com** `job_id`: `formData.append('job_id', getNcJobId())`. |
| Última etapa do fluxo | Incluir `formData.append('finalize', '1')`. |
| Download dos arquivos | Usar `download_urls` do JSON ou montar `GET /outputs/nc/{job_id}/{subpath}` (ex.: `final/nc_separados.zip`). |
| Stream (ZIP/XLSX) | Pedir `?format=zip` ou `?format=xlsx`; ler `X-NC-Job-Id` e chamar `setNcJobId(jobId)`. |

---

## 6. Ordem das etapas (referência)

1. **Separar** (M01) — cria job se não houver `job_id`
2. **Gerar modelo foto** (M02)
3. **Inserir conservação** (M03) ou **Inserir meio ambiente** (M07)
4. **Juntar** (M04)
5. **Inserir número** (M05) — típica última etapa; enviar `finalize=1` aqui se for o fim do fluxo

Alternativa em 2 chamadas: **POST /nc/start** (1 EAF) → **POST /nc/stage2** (job_id + params); o backend já marca finished e aplica retenção 72h.

---

## 7. Chave localStorage

- **Chave:** `artesp_nc_job_id`
- **Valor:** string do `job_id` (ex.: `nc_20260224_123456_abc123`).
- **Limpeza:** remover ao iniciar novo fluxo ou ao fazer logout, para não reutilizar job antigo por engano.
