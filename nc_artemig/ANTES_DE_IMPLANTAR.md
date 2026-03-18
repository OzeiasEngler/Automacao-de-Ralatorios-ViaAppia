# Antes de implantar (Artemig / lote 50)

Checklist após **adaptações de template** ou **alteração de mapeamento**. Objetivo: não subir produção com colunas ou ficheiro modelo desalinhados.

## 1. Template Kcor (Exportar Kcor)

| Item | Valor no projeto |
|------|------------------|
| Ficheiro | `nc_artemig/assets/Template/templates/_Planilha Modelo Kcor-Kria_artemig.xlsx` |
| Constante | `MODELO_KCOR_KRIA` em `nc_artemig/config.py` |
| Aba usada pelo Python | **Dados** (se não existir, usa `active`) |

**Após trocar o XLSX:** abrir o modelo, confirmar que a aba **Dados** tem na **linha 1** os cabeçalhos na mesma ordem que `COL_KCOR_KRIA` em `config.py` (A=NumItem … Y=Unidade). Se a equipa renomeou colunas ou inseriu colunas, **atualizar `COL_KCOR_KRIA`** (índices 1–25).

**Nota:** a macro Nas01 no Excel abre *Planilha Padrão Kcor - Nascentes.XLSX* (rede). No repositório a referência é o **Kcor-Kria** acima — deve ser a versão **equivalente** à que a CONSOL/Kria aceita na importação.

## 2. Template análise PDF (relatório Excel)

| Item | Valor |
|------|--------|
| Ficheiro | `nc_artemig/assets/Template/Template_EAF_artemig.xlsx` |
| Constante | `TEMPLATE_RELATORIO_ANALISE_PDF` em `config.py` |

Confirmar colunas esperadas pelo parser (Tipo, SH, Nº CONSOL, etc.) conforme última adaptação.

## 3. Mapeamentos a validar

| Área | Onde está | O que conferir |
|------|-----------|----------------|
| Colunas A–Y Kcor | `config.py` → `COL_KCOR_KRIA` | Índices = posição real no Excel |
| Patologia → col. E (Tipo) | `exportar_kcor_planilha.py` → `_patologia_para_kcor` | Comparar amostra com Nas01 / negócio |
| Sentido → col. I | `sentido_kcor.py` + uso em `analisar_pdf_nc` | MG 050 / BR 265 / BR 491 |
| Trechos / EAF lote 50 | `config.py` → `MAPA_EAF_POR_LOTE`, `RODOVIAS_POR_LOTE` | Km e nomes atualizados ao contrato |

## 3.1 Fotos — colunas V e W (Exportar Kcor)

- **V:** `{base}\{subpasta}` — base = **`ARTEMIG_KCOR_DIR_FOTOS`** (padrão `O:\NOTIFICAÇÕES DER\_LANÇAR\_02 - Arquivos Fotos`). Subpasta = stem do PDF **ou**, se faltar, `NOT-yy-xxxxx_PAVIMENTO_CE{nº_consol}` derivado do código (9 dígitos) + Nº Consol (ex.: `NOT-25-01365_PAVIMENTO_CE2516929`). Sem essa subpasta as fotos não coincidem com a ocorrência no disco.
- **W:** ficheiros **dentro dessa subpasta:** `PDF (COD).jpg`, `nc (COD).jpg`, `nc (COD)_1.jpg`, … — o **COD** é sempre o da **mesma linha** (código da fiscalização dessa NC). Linhas ordenadas por código numérico para cruzar com o relatório.

## 4. Teste mínimo sugerido

1. Gerar **Exportar Kcor** localmente (um PDF lote 50 ou script com `NcItem` lote `50`).
2. Abrir o XLSX gerado: conferir **uma linha por NC**, datas, prazo (24h emergencial → 1 dia se for a regra acordada), col. B/C/D/E coerentes com Nas01.
3. Se V/W forem preenchidos noutra fase (Nas02/fotos), confirmar que o deploy inclui esse fluxo ou que fica documentado como manual.

## 5. API / deploy

- Ficheiros em `nc_artemig/assets/` devem ir no pacote (não ignorar no `.dockerignore` / artefacto).
- Lote **50** no ZIP da análise PDF: terceiro ficheiro `*Exportar Kcor.xlsx` — smoke test no ambiente alvo.

---

*Última revisão: alinhar este ficheiro sempre que o modelo XLSX ou `COL_KCOR_KRIA` mudarem.*
