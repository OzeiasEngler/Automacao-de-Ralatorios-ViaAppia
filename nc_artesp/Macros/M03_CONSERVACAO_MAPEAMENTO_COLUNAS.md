# Módulo 03 – Conservação: mapeamento de colunas (macro Art_03)

Referência exata das informações que entram em cada coluna do Excel, conforme a macro **Art_03_Inserir_NC_Rot_Salva_apo** (03 - Art_03_Inserir_NC_Rot_Salva_apo.bas).

---

## 1. ORIGEM: Planilha Kria (Arquivo Foto - Conserva)

Arquivo lido: **Kria** (ex.: `yyyymmdd-hhmm - 20260221 - CONSTATAÇÕES NC LOTE 13 ...xlsx`).  
Bloco de **5 linhas** por NC; âncora da primeira NC em **y = 9** (depois y = 14, 19, …).

Para cada NC, a macro lê (variável `y` = linha do km, ex.: 9 para a 1ª NC):

| Coluna | Linha   | Variável macro   | Conteúdo (origem) |
|--------|--------|-------------------|-------------------|
| **B**  | y - 3  | numero            | Número sequencial da NC (1, 2, 3…) — usado no nome do JPG e formatado 00000/0000/000 |
| **G**  | y - 2  | embasamento(i)    | Data envio (no Kria novo M02 = data constatação) |
| **D**  | y - 1  | rodovia(i)        | Rodovia (ex.: SP-075, SP-127, SP-300…) |
| **F**  | y - 1  | Sentido(i)        | Sentido (ex.: Norte, Sul, Leste, Oeste) |
| **G**  | y - 1  | texto(i) / descricao | Descrição do serviço / tipo NC (ex.: "Limpeza e varredura...") → vira serv(i) e depois classifica/executor por tabela |
| **D**  | y      | kminicial_t(i)    | Km inicial (ex.: 10+500) |
| **F**  | y      | kmfinal_t(i)      | Km final |
| **H**  | y      | codigo(i)         | Código da NC |
| **L**  | y      | complemento(i)    | Complemento |
| **L**  | y      | foto(i)           | Número da foto (para pdf (N).jpg) |
| **F**  | y + 1  | data(i)           | Data vencimento / data reparo (dd/mm/aaaa) |
| **H**  | y + 1  | relatorio(i)      | Número do relatório EAF |
| **L**  | y + 1  | Prazo(i)          | Prazo em dias (número) |

Resumo por linha do bloco (exemplo y = 9):

- **Linha 6 (y-3):** B = sequência.
- **Linha 7 (y-2):** G = embasamento (data envio).
- **Linha 8 (y-1):** D = rodovia, F = sentido, G = texto/descrição (tipo NC).
- **Linha 9 (y):**   D = km inicial, F = km final, H = código, L = complemento, L = nº foto.
- **Linha 10 (y+1):** F = data (vencimento/reparo), H = relatório, L = prazo (dias).

A macro ainda normaliza **rodovia** (SP-075 → SP075, etc.) e converte **descrição** (texto) em **serv**, **classifica** e **executor** pela tabela de equivalências (Pichação, Pav. - Buraco, Drenagem - Limpeza, etc.).

---

## 2. DESTINO: Planilha Kcor-Kria (_Planilha Modelo Kcor-Kria.XLSX)

Modelo aberto pela macro; dados gravados a partir da **linha 2** (j = 2), uma linha por NC (x = 1, 2, 3…).

Cabeçalho do template (config: `CABECALHO_KCOR_KRIA`):

1. NumItem  
2. Origem  
3. Motivo  
4. Classificação  
5. Tipo  
6. Rodovia  
7. KMi  
8. KMf  
9. Sentido  
10. Local  
11. Gestor  
12. Executor  
13. Data Solicitação  
14. Data Suspensão  
15. DtInicio_Prog  
16. DtFim_Prog  
17. DtInicio_Exec  
18. DtFim_Exec  
19. Prazo  
20. ObsGestor  
21. Observações  
22. Diretório  
23. Arquivos  
24. Indicador  
25. Unidade  

O que a **macro** grava em cada coluna (exatamente como no VBA):

| Col | Nome no template      | Valor gravado pela macro |
|-----|------------------------|---------------------------|
| **A** | NumItem             | x (índice 1, 2, 3…) |
| **B** | Origem              | "Artesp" |
| **C** | Motivo              | "2" |
| **D** | Classificação       | classifica(x) — ex.: "Conservação Rotina", "Sinalização" |
| **E** | Tipo                | serv(x) — ex.: "Pav. - Limpeza", "Pichação" |
| **F** | Rodovia             | rodovia(x) — já normalizado (SP075, SP127, …) |
| **G** | KMi                 | kminicial_t(x) |
| **H** | KMf                 | kmfinal_t(x) |
| **I** | Sentido             | Sentido(x) |
| **J** | Local               | *(não escrito pela macro; fica em branco ou valor do modelo)* |
| **K** | Gestor              | "Conservação" |
| **L** | Executor            | executor(x) — ex.: "Soluciona - Conserva" |
| **M** | Data Solicitação    | Format(data(x), "mm/dd/yyyy") — **data = F(y+1) = data constatação (F10)** |
| **N** | Data Suspensão      | *(não escrito pela macro)* |
| **O** | DtInicio_Prog       | Format(data(x), "mm/dd/yyyy") — **mesmo que M (data constatação)** |
| **P** | DtFim_Prog          | Format(embasamento(x), "mm/dd/yyyy") — **embasamento = G(y-2) = data envio** |
| **Q** | DtInicio_Exec       | *(não escrito pela macro)* |
| **R** | DtFim_Exec          | *(não escrito pela macro)* |
| **S** | Prazo               | Prazo(x) — número (dias) |
| **T** | ObsGestor           | "--> Relatório EAF Conservação Rotina nº: " & relatorio(x) & vbCrLf & "--> Código NC: " & codigo(x) — **uma única quebra** entre relatório e código; relatorio = H(y+1), codigo = H(y) na entrada |
| **U** | Observações         | texto(x) & vbCrLf & vbCr & (se complemento vazio: "- Data Superação Artesp ----> " & embasamento; senão: "- Complemento ----> " & complemento(x) & vbCrLf & vbCrLf & vbCr & "- Embasamento ----> " & embasamento(x)) — **uma quebra** após texto; **duas quebras** entre Complemento e Embasamento |
| **V** | Diretório          | Diretorio — caminho da pasta de imagens (ex.: L:\...\Imagens\Conservação) |
| **W** | Arquivos            | arquivo(i) & ";" & "pdf (" & foto(i) & ").jpg" — nome do JPG gerado e foto PDF |

Colunas **J, N, Q, R** não são preenchidas pela macro (permanecem do modelo ou vazias).  
Coluna **Y** (Unidade) é preenchida depois pelo **Módulo 05** (Inserir Número Kria).

---

## 3. Resumo origem → destino (Kria → Kcor-Kria)

| Kcor-Kria (col) | Fonte no Kria |
|------------------|----------------|
| A NumItem        | B(y-3) sequência |
| B Origem         | fixo "Artesp" |
| C Motivo         | fixo "2" |
| D Classificação  | derivado de G(y-1) via tabela (ex.: "Conservação Rotina") |
| E Tipo           | derivado de G(y-1) via tabela (serv) |
| F Rodovia        | D(y-1) normalizado (SP075, SP127, …) |
| G KMi            | D(y) km inicial |
| H KMf            | F(y) km final |
| I Sentido        | F(y-1) |
| K Gestor         | fixo "Conservação" |
| L Executor       | derivado de G(y-1) via tabela |
| M Data Solicitação | F(y+1) data constatação (Format mm/dd/yyyy) |
| O DtInicio_Prog  | F(y+1) data constatação (igual M) |
| P DtFim_Prog     | G(y-2) embasamento / data envio (Format mm/dd/yyyy) |
| S Prazo          | L(y+1) prazo em dias |
| T ObsGestor      | H(y+1) relatório + H(y) código |
| U Observações    | G(y-1) texto + L(y) complemento + G(y-2) embasamento |
| V Diretório      | constante Diretorio (pasta Imagens\Conservacao) |
| W Arquivos       | nome do JPG exportado + ";pdf (foto).jpg" |

---

## 4. Implementação Python (inserir_nc_kria.py)

O mapeamento equivalente está em `inserir_nc_kria.py`: constantes `_K_A` … `_K_Y` e a função que monta a planilha Kcor-Kria (FASE 2). As colunas de data são detectadas pelo cabeçalho (`_detectar_colunas_data_kcor`) para compatibilidade com templates que usem "Data Envio" / "Data Solicitação" e "Data Reparo" / "DtFim_Prog".
