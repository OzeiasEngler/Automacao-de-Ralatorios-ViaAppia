# Mapa de Dependências — GeradorARTESP

Análise do que está **em uso** pelos módulos e da **raiz** (`assets/`, `nc_artesp/assets/`, `fotos_campo/assets/`).

---

## 1. `assets/` (raiz do projeto)

Usado por: `gerador_artesp_core.py`, `render_api/app.py`

| Caminho | Referência | Status |
|---------|------------|--------|
| `assets/schema/conserva.schema.r0.json` | `SCHEMA_PATH`, `mod["schema_asset"]` | ✅ Em uso |
| `assets/schema/obras.schema.r0.json` | `SCHEMA_PATH`, `mod["schema_asset"]` | ✅ Em uso |
| `assets/malha/Eixo Lote 13.xlsx` | `LOTES["13"]["eixo"]` | ✅ Em uso |
| `assets/malha/Eixo lote 21.xlsx` | `LOTES["21"]["eixo"]` | ✅ Em uso |
| `assets/malha/Eixo Lote 26.csv` | `LOTES["26"]["eixo"]` | ✅ Em uso |
| `assets/malha/Eixo Lote 13 1.csv` | — | ⚠️ Não referenciado no código |
| `assets/template/L13_conservacao_2026_r0.xlsx` | `_path_asset_template` | ✅ Em uso |
| `assets/template/L13_obras_2026_r0.xlsx` | `_path_asset_template` | ✅ Em uso |
| `assets/template/L21_*`, `L26_*` | idem | ✅ Em uso |

---

## 2. `nc_artesp/assets/`

Usado por: `nc_artesp/config.py`, `separar_nc.py`, `verificar_merge_fotos.py`, `render_api/fotos_router.py`, `render_api/nc_router.py`

O código procura em **`assets/templates`** primeiro e depois em **`assets/`** (fallback).

### Templates referenciados no código

| Arquivo | Quem usa | Caminho esperado |
|---------|----------|------------------|
| `Template_EAF.xlsx` ou `Template_EAF.xlsx.xlsx` | M01 Separar NC, M02 MA | `assets/Template` ou `assets/templates` |
| `Modelo Abertura Evento Kria Conserva Rotina.xlsx` | M02 Gerar Modelo Foto, M07 | `assets/templates` ou `assets/` |
| `Modelo.xlsx` | M02 Resposta | `assets/templates` ou `assets/` |
| `_Planilha Modelo Kcor-Kria.XLSX` | M03, M07 Inserir NC | `assets/templates` ou `assets/` |
| `Planilha Modelo Conservação - Foto 2 Lados.xlsx` | `fotos_router` (API Fotos) | `nc_artesp/assets/` |
| `Eventos Acumulado Artesp para Exportar Kria.xlsx` | M04 Juntar | `assets/templates` |

### Demais arquivos em `nc_artesp/assets/`

Usados pelas macros VBA/Excel e como referência de modelos (não referenciados diretamente no código Python, mas necessários para o fluxo NC):

- `20200504 - PERSONAL.XLSB` — macro Excel
- `Abertura de evento.XLSX`, `Check List.xlsx`, etc. — modelos das macros
- `_Planilha Modelo nc lote 13.xls` — mencionado em docstring como template manual
- Demais `.xls`, `.xlsx`, `.xlsm` — modelos e macros ARTESP

Status: **necessários** para execução das macros e do pipeline NC (desktop/API).

---

## 3. `fotos_campo/assets/`

Usado por: `render_api/fotos_router.py`

| Arquivo | Quem usa | Status |
|---------|----------|--------|
| `Relação Total - Lote 13.xlsx` | `_carregar_relacao_assets("13")` | ✅ Em uso |
| `Relação Total - Lote 21.xlsx` | `_carregar_relacao_assets("21")` | ✅ Em uso |
| `Relação Total - Lote 26.xlsx` | `_carregar_relacao_assets("26")` | ✅ Em uso |
| `Template/Planilha Modelo Conservação - Foto 2 Lados.xlsx` | — | ⚠️ `fotos_router` usa `nc_artesp/assets/` para esse template, não `fotos_campo/assets/` |

O `fotos_router` carrega o template Foto 2 Lados de **`nc_artesp/assets/`**, não de `fotos_campo/assets/Template/`. O arquivo em `fotos_campo` pode ser usado por `fotos_campo/core.py` em modo desktop ou como alternativa local.

---

## 4. `render_api/` — dependências de assets

- **Schema:** `assets/schema/*.json` (via `SCHEMA_PATH`)
- **Malha:** `assets/malha/` (via `gerador_artesp_core.LOTES`)
- **Templates conservação/obras:** `assets/template/` (via `_path_asset_template`)
- **NC:** `nc_artesp/assets/` ou `nc_artesp/assets/templates`
- **Fotos:** `fotos_campo/assets/` (Relação Total) e `nc_artesp/assets/` (template Foto 2 Lados)

---

## Resumo

| Pasta | Em uso | Não referenciado |
|-------|--------|------------------|
| `assets/schema/` | conserva, obras | — |
| `assets/malha/` | Eixo Lote 13.xlsx, Eixo lote 21.xlsx, Eixo Lote 26.csv | Eixo Lote 13 1.csv |
| `assets/template/` | L13/L21/L26 conservacao/obras | — |
| `nc_artesp/assets/` | vários modelos e templates | duplicatas (assets vs templates) |
| `fotos_campo/assets/` | Relação Total L13/L21/L26 | Template/ pode ser redundante (usa nc_artesp) |

---

## Nota sobre `nc_artesp/assets/templates/` vs `nc_artesp/assets/`

- **config.py** aponta para `assets/templates/`
- **nc_router** e **verificar_merge_fotos** usam fallback para `assets/` se `templates/` não existir
- Se `templates/` foi removido do repositório, o fallback em `assets/` continua funcionando enquanto os mesmos arquivos existirem em `assets/`
