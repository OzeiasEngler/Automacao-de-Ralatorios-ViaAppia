"""
modulos/juntar_arquivos.py
────────────────────────────────────────────────────────────────────────────
Equivalente VBA: Art_04_EAF_Rot_Juntar_Arquivo_Exportar_Kria
Desenvolvedor: Ozeias Engler

O Módulo Separar NC gera um .xlsx por NC; uma linha de dados por NC (há mais linhas só se houver duas NCs do mesmo tipo).
O acumulado junta todos em uma planilha:
  • Coluna A no acumulado = somente contagem de itens: 1 linha → item 1, 2 linhas → itens 1 e 2, etc.
  • B–Y da NC separada = mesma linha no acumulado (dados válidos).
"""

import logging
from pathlib import Path
from datetime import datetime

import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Border, Side

from config import M04_ENTRADA, M04_ACUMULADO, M04_SAIDA, M04_NOME_SAIDA, CABECALHO_KCOR_KRIA, NUM_COLUNAS_KCOR_KRIA
from utils.helpers import garantir_pasta, parse_data, data_yyyymmdd

logger = logging.getLogger(__name__)

# Número total de colunas da planilha Kcor-Kria (compatível VBA)
NUM_COLUNAS = NUM_COLUNAS_KCOR_KRIA


# Ordem canônica A–Y (mesma do config e da macro)
_CABECALHO_ORDEM = tuple(CABECALHO_KCOR_KRIA)  # 25 nomes
# Coluna M = Data Solicitação (índice 12 em 0-based)
_COL_DATA_SOLICITACAO = 13

_SIDE_THIN = Side(style="thin", color="000000")
_BORDA_PADRAO = Border(
    left=_SIDE_THIN, right=_SIDE_THIN, top=_SIDE_THIN, bottom=_SIDE_THIN
)


def _normalizar_header(s: str) -> str:
    """Normaliza nome de coluna para comparação (minúsculo, sem acentos)."""
    if s is None:
        return ""
    s = str(s).strip().lower()
    for old, new in [("ã", "a"), ("á", "a"), ("ç", "c"), ("é", "e"), ("ê", "e"), ("í", "i"), ("ó", "o"), ("ô", "o"), ("ú", "u")]:
        s = s.replace(old, new)
    return s


def _mapear_colunas_pelo_cabecalho(ws) -> list[int]:
    """
    Entrada: ws (planilha), linha 1 = cabeçalhos.
    Saída: lista de 25 int (1-based), índice i = coluna física do i-ésimo cabeçalho canônico (A=0..Y=24).
    Fallback: se cabeçalho não encontrado, usa posição i+1. Aliases: Executor/Executores, Data Envio/Data Solicitação, Arquivo/Arquivos.
    Porquê: arquivos de entrada (M03) podem ter colunas em ordem diferente ou mescladas; saída do M04 deve ser sempre ordem canônica A–Y.
    """
    mapa = {}  # nome_normalizado -> col (1-based)
    for c in range(1, min(ws.max_column + 1, 50)):
        v = _valor_celula(ws, 1, c, preencher_se_merge=True)
        if v is None:
            continue
        n = _normalizar_header(str(v))
        if n and n not in mapa:
            mapa[n] = c
    # Aliases: template pode ter "Executores"/"Executor", "Data Envio"/"Data Solicitação", "Arquivo"/"Arquivos"
    for n, col in list(mapa.items()):
        if n == "executores" and "executor" not in mapa:
            mapa["executor"] = col
        if n == "executor" and "executores" not in mapa:
            mapa["executores"] = col
        if n == "data envio" and "data solicitacao" not in mapa:
            mapa["data solicitacao"] = col
        if n == "data solicitacao" and "data envio" not in mapa:
            mapa["data envio"] = col
        if n == "arquivo" and "arquivos" not in mapa:
            mapa["arquivos"] = col
        if n == "arquivos" and "arquivo" not in mapa:
            mapa["arquivo"] = col
    out = []
    for i, nome in enumerate(_CABECALHO_ORDEM):
        n = _normalizar_header(nome)
        col = mapa.get(n)
        if col is None:
            col = i + 1
        out.append(col)
    return out


def _nome_saida_macro(todos_registros: list, nome_base: str = M04_NOME_SAIDA) -> str:
    """
    Nome do arquivo de saída igual à macro Art_04 (linhas 224-229):
      dia = Left(Data_Solicitação(g - 1), 2)
      mes = Right(Left(Data_Solicitação(g - 1), 5), 2)
      ano = Right(Left(Data_Solicitação(g - 1), 10), 4)
      NameFile = ano & mes & dia & " - " & Format(Now, "hhmmss") & " - Eventos Acumulado Artesp para Exportar Kria.xlsx"
    Ou seja: YYYYMMDD (da data do último registro) - hhmmss (hora atual) - nome_base
    """
    ano, mes, dia = None, None, None
    if todos_registros:
        ultimo = todos_registros[-1]
        if len(ultimo) >= _COL_DATA_SOLICITACAO and ultimo[_COL_DATA_SOLICITACAO - 1] is not None:
            s = str(ultimo[_COL_DATA_SOLICITACAO - 1]).strip()
            # Macro: dia=Left(s,2), mes=Right(Left(s,5),2), ano=Right(Left(s,10),4) → formato DD/MM/YYYY ou DD-MM-YYYY
            if len(s) >= 10:
                dia = s[:2]
                mes = s[3:5]
                ano = s[6:10]
    if ano is None or mes is None or dia is None:
        now = datetime.now()
        ano = now.strftime("%Y")
        mes = now.strftime("%m")
        dia = now.strftime("%d")
    hhmmss = datetime.now().strftime("%H%M%S")
    return f"{ano}{mes}{dia} - {hhmmss} - {nome_base}"


def criar_base_acumulado(caminho: Path) -> None:
    """Planilha acumulada mínima (só cabeçalho); usada quando não há base enviada."""
    wb = openpyxl.Workbook()
    ws = wb.active
    for c, val in enumerate(CABECALHO_KCOR_KRIA, start=1):
        ws.cell(row=1, column=c).value = val
    caminho.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(caminho))
    wb.close()


def _valor_celula(ws, row: int, col: int, preencher_se_merge: bool = False):
    """Valor da célula. Em merge, openpyxl retorna None nas células que não são o canto superior-esquerdo.
    preencher_se_merge=True para W,X,Y: arquivos M03 podem ter essas colunas mescladas; sem isso o acumulado ficaria vazio em 22–25."""
    for merged_range in ws.merged_cells.ranges:
        if row < merged_range.min_row or row > merged_range.max_row:
            continue
        if col < merged_range.min_col or col > merged_range.max_col:
            continue
        if row != merged_range.min_row or col != merged_range.min_col:
            if preencher_se_merge:
                return ws.cell(row=merged_range.min_row, column=merged_range.min_col).value
            return None
        break
    return ws.cell(row=row, column=col).value


def _aplicar_bordas_linha(ws, row: int, col_fim: int = NUM_COLUNAS):
    """Borda em células da linha 1..col_fim."""
    for col in range(1, col_fim + 1):
        ws.cell(row=row, column=col).border = _BORDA_PADRAO


def _copiar_bordas_linha(ws, row_origem: int, row_destino: int, col_fim: int = NUM_COLUNAS):
    """Copia borda da linha origem para a destino (preserva formatação do template)."""
    for col in range(1, col_fim + 1):
        src = ws.cell(row=row_origem, column=col)
        dst = ws.cell(row=row_destino, column=col)
        if src.border and getattr(src.border, "left", None) is not None:
            dst.border = src.border.copy()
        else:
            dst.border = _BORDA_PADRAO


def _celula_preenchida(val) -> bool:
    """True se o valor existe e não é string vazia (para detectar última linha em A)."""
    if val is None:
        return False
    if isinstance(val, str) and not val.strip():
        return False
    return True


def _ultima_linha_col_a(ws, max_row: int) -> int:
    """
    Última linha com dado na coluna A, igual à macro: Cells(65536, 1).End(xlUp).Row.
    """
    for r in range(max_row, 0, -1):
        if _celula_preenchida(ws.cell(row=r, column=1).value):
            return r
    return 1


def _ultima_linha_abc(ws, max_row: int) -> int:
    """Última linha com dado em A, B ou C (fallback quando A vem vazio, ex.: fórmulas com data_only=True)."""
    for r in range(max_row, 0, -1):
        for col in (1, 2, 3):
            if _celula_preenchida(ws.cell(row=r, column=col).value):
                return r
    return 1


def _eh_cabecalho(linha: list) -> bool:
    """True se a linha for o cabeçalho (A = 'NumItem')."""
    if not linha:
        return False
    a = linha[0]
    return a is not None and str(a).strip().upper() == "NUMITEM"


def _ultima_linha_qualquer_col(ws, max_row: int, colunas: tuple = (1, 2, 3, 4, 5)) -> int:
    """Última linha com dado em qualquer uma das colunas (ex.: A–E)."""
    for r in range(max_row, 0, -1):
        for col in colunas:
            if _celula_preenchida(ws.cell(row=r, column=col).value):
                return r
    return 1


def _obter_planilha_e_ultima(wb):
    """
    Retorna (ws, ultima): planilha e última linha a ler.
    Igual à macro: ultimalinhaprov = Cells(65536, 1).End(xlUp).Row — só coluna A.
    Se ativa der ultima<=1, tenta aba 'Dados'; se ainda 1, usa max_row.
    """
    ws = wb.active
    max_row = ws.max_row
    ultima = _ultima_linha_col_a(ws, max_row)
    if ultima <= 1 and len(wb.worksheets) > 1:
        for sheet in wb.worksheets:
            if sheet.title and "dados" in sheet.title.lower():
                mr = sheet.max_row
                u = _ultima_linha_col_a(sheet, mr)
                if u > 1:
                    return sheet, u
                if mr >= 2:
                    return sheet, mr
    if ultima <= 1 and max_row >= 2:
        ultima = max_row
    return ws, ultima


def _ler_arquivo(caminho: Path) -> list[list]:
    """
    Lê .xlsx: linhas 2 até última com dado na coluna A.
    Colunas são lidas pelo CABEÇALHO (linha 1), não pela posição fixa, para não
    remontar dados em colunas erradas quando o arquivo tem ordem diferente ou mescladas.
    Cada linha vira um registro na ordem canônica A–Y (25 colunas).
    """
    wb = load_workbook(str(caminho), data_only=True)
    ws, ultima = _obter_planilha_e_ultima(wb)
    col_map = _mapear_colunas_pelo_cabecalho(ws)  # [col_A, col_B, ...] 1-based

    linhas = []
    for r in range(2, ultima + 1):
        # Ler cada valor na coluna correta pelo nome do cabeçalho (ordem canônica)
        # W,X,Y: preencher_se_merge (fontes M03 podem ter merge)
        linha = []
        for i in range(NUM_COLUNAS):
            col = col_map[i] if i < len(col_map) else (i + 1)
            preencher = i >= 22  # Arquivos, Indicador, Unidade
            linha.append(_valor_celula(ws, r, col, preencher_se_merge=preencher))
        if _eh_cabecalho(linha):
            continue
        while len(linha) < NUM_COLUNAS:
            linha.append(None)
        linhas.append(linha[:NUM_COLUNAS])
    wb.close()
    return linhas


def executar(pasta_entrada: Path | None = None,
             arquivo_acumulado: Path | None = None,
             pasta_saida: Path | None = None,
             nome_saida: str | None = None,
             nome_arquivo_completo: str | None = None,
             callback_progresso=None,
             arquivos_entrada: list[Path] | None = None) -> Path | None:
    """
    M04 Juntar: consolida .xlsx Kcor-Kria individuais numa planilha acumulada (uma linha por registro, A–Y).
    Entrada: pasta_entrada (ou arquivos_entrada) com .xlsx; arquivo_acumulado = base existente (cabeçalho + dados); pasta_saida.
    Saída: Path do .xlsx gerado em pasta_saida. nome_arquivo_completo sobrescreve nome gerado por data.
    Retorno None: nenhum .xlsx em entrada, ou arquivo_acumulado não existe (quando obrigatório).
    """
    pasta_entrada     = pasta_entrada     or M04_ENTRADA
    arquivo_acumulado = arquivo_acumulado or M04_ACUMULADO
    pasta_saida       = pasta_saida       or M04_SAIDA
    garantir_pasta(pasta_saida)

    if arquivos_entrada is not None:
        arquivos = sorted([
            Path(f) for f in arquivos_entrada
            if Path(f).exists()
            and Path(f).suffix.lower() == ".xlsx"
            and not Path(f).name.startswith("~")
            and "Acumulado" not in Path(f).name
            and not Path(f).name.startswith("_")
        ])
    else:
        arquivos = sorted([
            f for f in pasta_entrada.glob("*.xlsx")
            if not f.name.startswith("~")
            and "Acumulado" not in f.name
            and not f.name.startswith("_")
        ])

    if not arquivos:
        logger.warning(f"Nenhum .xlsx encontrado em: {pasta_entrada}")
        return None

    logger.info(f"Encontrados {len(arquivos)} arquivo(s) para consolidar.")
    todos_registros: list[list] = []
    for idx, arq in enumerate(arquivos):
        if callback_progresso:
            callback_progresso(idx + 1, len(arquivos), f"Lendo: {arq.name[:60]}")
        logger.info(f"Lendo: {arq.name}")
        registros = _ler_arquivo(arq)
        if not registros:
            logger.warning("  %s: 0 registro(s) (verifique se a planilha tem dados na linha 2+, coluna A ou B/C)", arq.name)
        todos_registros.extend(registros)
        logger.info(f"  {len(registros)} registro(s) lido(s).")

    if not todos_registros:
        logger.warning("Nenhum registro encontrado nos arquivos.")
        return None

    logger.info(f"Total de registros a consolidar: {len(todos_registros)}")
    if not arquivo_acumulado.exists():
        logger.warning("Acumulado não informado. Envie o arquivo acumulado (relatório da rede) para consolidar.")
        return None

    wb_acum = load_workbook(str(arquivo_acumulado))
    ws_acum = None
    for sheet in wb_acum.worksheets:
        a1 = sheet.cell(row=1, column=1).value
        if a1 is not None and "numitem" in str(a1).strip().lower():
            ws_acum = sheet
            break
    if ws_acum is None:
        ws_acum = wb_acum.worksheets[0]
        logger.debug("Usando primeira planilha (cabeçalho 'NumItem' não encontrado em A1).")

    logger.info(
        f"Acumulado: planilha '{ws_acum.title}'. "
        f"Gravando {len(todos_registros)} registro(s) nas colunas A–Y a partir da linha 2."
    )
    max_row_acum = ws_acum.max_row
    N = len(todos_registros)

    for idx, registro in enumerate(todos_registros):
        row = 2 + idx
        ws_acum.cell(row=row, column=1).value = idx + 1  # A = contagem
        for col in range(2, NUM_COLUNAS + 1):
            val = registro[col - 1] if (col - 1) < len(registro) else None
            ws_acum.cell(row=row, column=col).value = val
    for row in range(2, 2 + N):
        _aplicar_bordas_linha(ws_acum, row)
    for r in range(2 + N, max_row_acum + 1):
        for c in range(1, NUM_COLUNAS + 1):
            ws_acum.cell(row=r, column=c).value = None
        _aplicar_bordas_linha(ws_acum, r)

    # Nome saída: macro Art_04 (YYYYMMDD - hhmmss - Eventos Acumulado...)
    if nome_arquivo_completo and nome_arquivo_completo.strip():
        nome_arq_saida = nome_arquivo_completo.strip()
        if not nome_arq_saida.lower().endswith(".xlsx"):
            nome_arq_saida += ".xlsx"
    else:
        nome_base = nome_saida if nome_saida else M04_NOME_SAIDA
        nome_arq_saida = _nome_saida_macro(todos_registros, nome_base)
    destino = pasta_saida / nome_arq_saida
    garantir_pasta(destino.parent)

    wb_acum.active = ws_acum
    wb_acum.save(str(destino))
    wb_acum.close()
    logger.info(f"Módulo 04 concluído. Acumulado salvo: {destino.name}")

    if callback_progresso:
        callback_progresso(len(arquivos), len(arquivos), "Módulo 04 concluído.")

    return destino
