# Lista para Push — GeradorARTESP

## Remote

```
origin  https://github.com/gestao-rodovias/Automacao-de-Ralatorios-ViaAppia.git
```

---

## Conteúdo já commitado (187 arquivos)

| Categoria | Arquivos/Pastas |
|-----------|-----------------|
| **Config** | .gitignore, .gitattributes, .dockerignore, Dockerfile, requirements.txt, run_local.bat |
| **Raiz** | README.md, gerador_artesp_core.py |
| **assets/schema** | conserva.schema.r0.json, obras.schema.r0.json, LEIA-ME.txt |
| **assets/malha** | Eixo Lote 13.xlsx, Eixo lote 21.xlsx, Eixo Lote 26.csv, Eixo Lote 13 1.csv |
| **assets/template** | L13/L21/L26 conservacao/obras 2026_r0.xlsx, LEIA-ME.txt, PLANILHA_L21_GEOJSON.md |
| **fotos_campo** | core.py, __init__.py, assets (Relação Total L13/21/26, Template Foto 2 Lados), Macros |
| **nc_artesp** | config.py, pdf_extractor.py, modulos/, utils/, Macros/, assets/, assets/templates/ |
| **render_api** | app.py, routers, auth_crypto, conformidade, plano_anual, job_manager, tests, web/, web-static/ |
| **Outros** | check_geojson_exe.py, exemplo_inserir_imagem_preenchendo_celula.py, preparar_push.ps1, docs .md |

---

## Arquivos não rastreados (opcionais)

| Arquivo | Descrição |
|---------|-----------|
| MAPA_DE_DEPENDENCIAS.md | Mapeamento de dependências |
| VERIFICACAO_ASSETS.md | Verificação de assets rastreados |

---

## Comandos

### Opção 1 — Push do que já está commitado

```powershell
cd C:\GeradorARTESP
git push -u origin main
```

### Opção 2 — Incluir os docs de mapeamento e verificação antes do push

```powershell
cd C:\GeradorARTESP
git add MAPA_DE_DEPENDENCIAS.md VERIFICACAO_ASSETS.md LISTA_PUSH.md
git commit -m "Docs: mapa de dependencias e verificacao de assets"
git push -u origin main
```

---

## Observação

Certifique-se de estar logado na conta **OzeiasEngler** no GitHub (ou use o GitHub Desktop). Se aparecer "Repository not found", confira credenciais ou crie o repositório em https://github.com/new.
