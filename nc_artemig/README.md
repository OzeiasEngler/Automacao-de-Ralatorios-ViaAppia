# nc_artemig — Pipeline NC Artemig (MG)

Pasta do regime **Artemig (MG)**. Contém assets, malhas e templates usados pelo pipeline de análise de NC quando `regime=artemig` (espelho do NC ARTESP para o estado de Minas Gerais).

## Config

Em **`config.py`** estão definidos:

- **Lote 50 — Nascentes das Gerais**: concessionária e rodovias **MG 050**, **BR 265**, **BR 491**.
- `LOTE_CONCESSIONARIA`, `RODOVIAS_POR_LOTE`, `LOTES_MENU_ANALISE`.
- `MAPA_EAF_POR_LOTE` e `MAPA_RESPONSAVEL_TECNICO_POR_LOTE` (a preencher quando houver EAFs/trechos por lote MG).

## Estrutura

| Pasta      | Uso |
|-----------|-----|
| **assets/**  | Modelos de planilha (Template EAF, modelo Kcor-Kria, Foto 2 Lados, etc.). Subpastas: `Template/`, `Malha/` conforme uso das macros/exportação. |
| **malha/**   | Malhas e eixos por lote (dados de referência para trechos, km, concessionárias MG). |
| **templates/** | Modelos de **relatório** do pipeline: template XLSX do relatório de análise e, se houver, modelo de relatório PDF específico Artemig. |

## Uso no pipeline

- **Análise PDF (lote 50):** o ZIP de saída inclui, além do PDF de análise e do `Relatorio_Fiscalizacao_*.xlsx`, o ficheiro **`yyyyMMdd-HHmm - Exportar Kcor.xlsx`** (planilha Nas01 / Kria CONSOL), gerado por `exportar_kcor_planilha.py`. Colunas V/W (fotos) ficam vazias até o fluxo Nas02 ou preenchimento manual.
- O **config** do regime Artemig (lotes MG, MAPA_EAF, MAPA_RESPONSAVEL_TECNICO) apontará os paths para `nc_artemig/assets/` e `nc_artemig/templates/`.
- Os **PDFs** de constatação Artemig podem ter layout diferente do ARTESP; o parser (em `nc_artesp` ou aqui) será escolhido conforme o regime.
- O plano de implementação está em **`PLANO_ARTEMIG_MG.md`** na raiz do projeto.
- **Antes de implantar** (template adaptado, colunas, mapeamento): ver **`ANTES_DE_IMPLANTAR.md`**.

## Observação

Arquivos de trabalho (PDFs de exemplo, exportações Kcor, ZIPs) podem ficar na raiz de `nc_artemig/` durante o desenvolvimento; o pipeline em produção usará apenas `assets/`, `malha/` e `templates/`.
