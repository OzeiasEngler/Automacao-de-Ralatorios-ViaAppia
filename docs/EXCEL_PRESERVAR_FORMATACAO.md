# Preservar bordas e formatação em `.xlsx`

## Por que some formatação?

- **`DataFrame.to_excel('arquivo.xlsx')`** gera um ficheiro **novo** (folha “limpa”). Bordas, cores e merges anteriores **não são herdados**.
- O mesmo vale para regravar o ficheiro inteiro sem carregar o workbook existente.

## Aberto recomendado: openpyxl em modo edição

1. `wb = openpyxl.load_workbook(caminho)` — carregar o ficheiro **existente**.
2. `ws = wb['Dados']` ou `wb.active` — folha correta.
3. Alterar **só o valor**: `ws.cell(row=r, column=c, value="...")` ou `.value = ...` — em princípio mantém estilo da célula (desde que não apagues linhas nem substituas a folha inteira).
4. `wb.save(caminho)` — gravar.

**Cuidados neste projeto:** inserir/apagar linhas, limpar ranges ou copiar estilo de uma linha modelo pode **mudar** bordas; por isso o pipeline NC (relatório, Kcor-Kria, Exportar Kcor) usa **cópia explícita de estilo** e/ou **reaplicar bordas** onde o template exige.

## Pandas + template existente

Se precisares de despejar um DataFrame **em cima** de um layout já formatado:

```python
import pandas as pd
from openpyxl import load_workbook

book = load_workbook(nome_arquivo)
with pd.ExcelWriter(
    nome_arquivo,
    engine="openpyxl",
    mode="a",
    if_sheet_exists="overlay",
) as writer:
    writer.book = book
    df.to_excel(writer, sheet_name="Planilha1", index=False, startrow=..., startcol=...)
```

`if_sheet_exists='overlay'` evita recriar a aba inteira; mesmo assim convém **backup** com gráficos, tabelas dinâmicas ou metadados complexos.

## Onde está o quê neste repositório

| Área | Abordagem |
|------|-----------|
| Relatório fiscalização / Kcor / templates NC | `openpyxl.load_workbook`, escrita por célula; ver `nc_artesp/modulos/analisar_pdf_nc.py`, `inserir_nc_kria.py`, `nc_artemig/exportar_kcor_planilha.py` |
| Export simples sem template | Ex.: `render_api/conformidade.py` — `to_excel` para ficheiro **novo** de auditoria (não é edição de modelo) |

Para utilitários (xls→xlsx, etc.): `nc_artesp/utils/excel_io.py`.
