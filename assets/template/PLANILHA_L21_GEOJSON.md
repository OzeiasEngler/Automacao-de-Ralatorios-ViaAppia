# Planilha L21 — Programação Anual para geração GeoJSON

Planilha de referência: **L21 - Programação Anual 2026.xlsx**  
(Ex.: `RELATÓRIOS ARTESP\RODOVIAS DO TIETE Lote 21\`)

## Estrutura esperada pelo gerador

- **Linha 1:** Cabeçalho (nomes das colunas).
- **Linhas 2–5:** Podem conter instruções ou exemplos (são ignoradas).
- **Linhas 6 em diante:** Dados (uma linha por trecho).

## Colunas reconhecidas

O gerador identifica as colunas pelo **nome normalizado** (sem acentos, espaços, etc.). Exemplos:

| Nome na planilha (exemplo)     | Nome interno      | Obrigatório |
|-------------------------------|-------------------|-------------|
| Lote                          | lote              | Sim         |
| Rodovia                       | rodovia           | Sim         |
| Item                          | item              | Sim         |
| Detalhamento / Serviço        | detalhamento_servico | Sim     |
| Unidade                       | unidade           | Sim         |
| Quantidade                    | quantidade        | Sim         |
| KM Inicial                    | km_inicial        | Sim         |
| KM Final                      | km_final          | Sim         |
| Local / Pista                 | local             | Sim         |
| Data Inicial                  | data_inicial      | Sim         |
| Data Final                    | data_final        | Sim         |
| Observações                   | observacoes_gerais| Sim (pode ser vazio) |
| Programa                      | programa          | Só Obras    |
| Subitem                       | subitem           | Só Obras    |
| Latitude / Longitude          | Latitude, Longitude | Opcional  |

## Lote 21

- A coluna **Lote** deve conter **L21** (ou "21", "Lote 21") para as linhas deste relatório.
- **Local:** use códigos do schema, com espaço ou underscore (ex.: Pista Norte, Marginal Sul, Canteiro Central). O gerador normaliza para PISTA_NORTE, MARGINAL_SUL, CANTEIRO_CENTRAL.
- **Sentidos:** Norte, Sul, Leste, Oeste são mapeados para a malha do L21 (Crescente/Decrescente).

## Verificar a planilha

Na pasta do projeto, execute:

```bash
python render_api/inspect_xlsx.py "C:\caminho\completo\L21 - Programação Anual 2026.xlsx"
```

O script mostra as colunas encontradas, quais foram reconhecidas pelo gerador e uma amostra dos dados.

## Gerar GeoJSON

1. Abra a API (ex.: `/web` ou `/web/geojson`).
2. Selecione **Lote 21**, **Conservação**, **R0**, **Ano 2026**.
3. Envie a planilha **L21 - Programação Anual 2026.xlsx**.
4. Baixe o GeoJSON e, se houver, o CSV de pendências (com resumo por motivo na primeira linha).
