# Render — Usar o repositório correto (GeradorARTESP)

Se a página inicial que abre no Render **não é deste projeto** (outra interface, outro título), o serviço está ligado ao **repositório errado** ou o Git local estava apontando para o projeto antigo.

## 1. Git local — apontar para GeradorARTESP

**Já corrigido:** o `origin` foi alterado para o repositório **GeradorARTESP**.

Para conferir:
```bash
git remote -v
```
Deve mostrar: `origin  https://github.com/oseiasengler/GeradorARTESP.git`

Se ainda aparecer `artesp-geojson-generator`, corrija com:
```bash
git remote set-url origin https://github.com/oseiasengler/GeradorARTESP.git
```

Assim, ao dar `git push`, o código sobe para o **GeradorARTESP**, e o Render (quando conectado a esse repo) fará o deploy correto.

## 2. Como corrigir no Render (painel)

1. Acesse **[dashboard.render.com](https://dashboard.render.com)** e entre no seu usuário.
2. Abra o serviço **artesp-geojson-api** (ou o nome que você deu).
3. Vá em **Settings** (Configurações).
4. Na seção **Build & Deploy** (ou **Repository**), localize **Repository** / **Connected Repository**.
5. Clique em **Change repository** (ou equivalente) e selecione:
   - **GeradorARTESP** (repositório que contém este projeto: GeoJSON, NC, Fotos, malhas L13/L21/L26).
6. Confirme. O **Root Directory** deve ficar **vazio** (raiz do repositório).
7. Faça um **Manual Deploy** (Deploy → Deploy latest commit) para subir o código do GeradorARTESP.

**Importante:** No Render, em **Settings** → **Repository**, o repositório conectado deve ser **GeradorARTESP** (não artesp-geojson-generator). Se estiver o antigo, use **Change repository** e selecione **GeradorARTESP**.

Depois do deploy, ao abrir a URL do serviço você deve ver:
- Título **"Gerador ARTESP — GeoJSON / Relatórios"**
- Texto **"Projeto: Conservação e Obras — GeoJSON e relatórios"**
- Redirecionamento para a página do **Gerador GeoJSON**.

Se ainda aparecer outra página, o serviço continua usando o repositório antigo — repita os passos e confirme que o repositório selecionado é **GeradorARTESP**.
