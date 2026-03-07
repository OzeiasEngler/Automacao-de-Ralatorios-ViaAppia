# Render — Usar o novo repositório (Automacao-de-Ralatorios-ViaAppia)

## Trocar repositório no serviço existente

1. Acesse **[dashboard.render.com](https://dashboard.render.com)** e faça login.
2. Abra o serviço **artesp-geojson-api** (ou o nome atual).
3. Vá em **Settings** → **Build & Deploy** → **Repository**.
4. Clique em **Change repository** e selecione:
   - **OzeiasEngler / Automacao-de-Ralatorios-ViaAppia**
5. Confirme. **Root Directory** deve ficar vazio (raiz do repositório).
6. Vá em **Manual Deploy** → **Deploy latest commit**.

---

## Ou criar um serviço novo (Blueprint)

1. Em **Dashboard** → **New** → **Blueprint**.
2. Conecte o GitHub e selecione **OzeiasEngler/Automacao-de-Ralatorios-ViaAppia**.
3. O Render detecta o `render.yaml` (ou `render_api/render.yaml`).
4. Se pedir **Root Directory**, deixe vazio ou use `.` (raiz).
5. Ajuste as variáveis de ambiente e clique em **Apply**.
6. Se houver serviço antigo, pode **pausar** ou **excluir** após o novo estar ok.

---

## Variáveis de ambiente (Settings → Environment)

| Chave | Descrição |
|-------|-----------|
| `ARTESP_JWT_SECRET` | Segredo JWT (produção) |
| `ARTESP_WEB_USERS` | Usuários: `user1:hash1,user2:hash2` |
| `ARTESP_WEB_PBKDF2_ITERATIONS` | Iterações PBKDF2 (opcional) |
| `ARTESP_PFX_PASSWORD` | Senha do certificado (opcional) |
| `ARTESP_PFX_CONTENT` | Certificado em Base64 (opcional) |

---

## URL do repositório

```
https://github.com/OzeiasEngler/Automacao-de-Ralatorios-ViaAppia
```
