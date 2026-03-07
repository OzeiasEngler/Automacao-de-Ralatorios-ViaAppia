# Verificação de Assets — GeradorARTESP

Verificação realizada: assets rastreados no git vs dependências nos scripts.

---

## ✅ Assets rastreados (ok)

### assets/ (raiz)
- schema: conserva.schema.r0.json, obras.schema.r0.json, LEIA-ME.txt
- malha: Eixo Lote 13.xlsx, Eixo lote 21.xlsx, Eixo Lote 26.csv, Eixo Lote 13 1.csv
- template: L13/L21/L26 conservacao/obras 2026_r0.xlsx

### fotos_campo/assets/
- Relação Total - Lote 13.xlsx
- Relação Total - Lote 21.xlsx
- Relação Total - Lote 26.xlsx
- Template/Planilha Modelo Conservação - Foto 2 Lados.xlsx

### nc_artesp/assets/
- Modelo Abertura Evento Kria Conserva Rotina.xlsx
- Modelo.xlsx
- _Planilha Modelo Kcor-Kria.XLSX
- Planilha Modelo Conservação - Foto 2 Lados.xlsx
- Template_EAF.xlsx.xlsx
- Demais modelos e macros (.xls, .xlsx, .xlsm, .xlsb)

### nc_artesp/assets/templates/
- Mesmos arquivos acima (duplicata para config que aponta em assets/templates/ primeiro)

---

## ⚠️ Possível ausência

| Arquivo | Referência | Status |
|---------|------------|--------|
| `nc_artesp/assets/templates/Template_EAF.xlsx` | config M01_TEMPLATE_EAF | Config espera .xlsx; no repo há `Template_EAF.xlsx.xlsx` e `Template_EAF.xls` |
| `nc_artesp/assets/templates/Eventos Acumulado Artesp para Exportar Kria.xlsx` | config M04_MODELO_ACUMULADO | Não está no repo; M04_MODELO_ACUMULADO não é referenciado em outro arquivo (só definido) |

---

## Conclusão

- **assets/ (raiz), fotos_campo/assets/, nc_artesp/assets/, nc_artesp/assets/templates/** estão rastreados.
- **Template_EAF**: separar_nc aceita `Template_EAF.xlsx` ou `Template_EAF.xlsx.xlsx`; há `.xlsx.xlsx` e `.xls` — falta apenas `Template_EAF.xlsx` se config apontar exatamente para esse nome. O separar_nc procura em várias pastas e aceita ambos os nomes.
- **Eventos Acumulado**: não está no repositório; o uso de M04_MODELO_ACUMULADO não foi encontrado no código além da definição em config.
