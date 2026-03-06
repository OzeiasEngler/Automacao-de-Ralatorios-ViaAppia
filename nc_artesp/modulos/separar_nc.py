"""
modulos/01_separar_nc.py
────────────────────────────────────────────────────────────────────────────
Equivalente VBA: Art_011_EAF_Separar_Mod_Exc_NC
Desenvolvedor: Ozeias Engler

A partir da planilha-mãe EAF (única, com todas as NCs do período),
gera arquivos XLS individuais — um por combinação Data+Rodovia+TipoNC.

Fluxo:
  1. Padroniza cols I e K para 3 dígitos (metros).
  2. Numera linhas na coluna V (sequencial de foto).
  3. Para cada linha, monta nome do arquivo destino.
  4. Para cada grupo: usa o template assets/templates/Template_EAF.xlsx (cópia binária);
     preenche com as linhas do grupo vindas da planilha-mãe (a partir da linha 5).
"""

import logging
import shutil
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl import load_workbook
import xlrd

from datetime import timedelta

from config import (
    M01_EXPORTAR,
    M01_LINHA_INICIO,
    M01_LOTE,
    M01_TEMPLATE_EAF,
    PRAZO_DIAS_APOS_ENVIO,
    RODOVIA_NOME_SEPARAR,
    SERVICO_ABREV,
    RODOVIAS,
)
from utils.helpers import (
    pad_metros,
    parse_data,
    data_yyyymmdd,
    data_ddmmaaaa,
    normalizar_rodovia_eaf,
    garantir_pasta,
    encurtar_nome_em_pasta,
    sanitizar_nome,
)

logger = logging.getLogger(__name__)


# ─────────────────────────────────────────────────────────────────────────────
# TEMPLATE EAF — planilha base para os arquivos gerados (cabeçalho 1–4; dados a partir da 5)
# Procura DENTRO do projeto nc_artesp (assets/Template ou assets/templates).
# Ordem: ARTESP_TEMPLATE_EAF (se definido) → nc_artesp/assets/Template → nc_artesp/assets/templates
# Aceita: Template_EAF.xlsx ou Template_EAF.xlsx.xlsx
# ─────────────────────────────────────────────────────────────────────────────
def _caminho_template_eaf() -> Path:
    """Retorna o Path do template EAF. Procura em nc_artesp/assets/ (dentro do projeto)."""
    # 1. ARTESP_TEMPLATE_EAF (ficheiro ou pasta)
    if M01_TEMPLATE_EAF.is_file():
        return M01_TEMPLATE_EAF
    if M01_TEMPLATE_EAF.is_dir():
        for n in ("Template_EAF.xlsx", "Template_EAF.xlsx.xlsx"):
            c = M01_TEMPLATE_EAF / n
            if c.is_file():
                return c
    # 2. Pastas dentro do pacote nc_artesp (projeto GeradorARTESP)
    _nc = Path(__file__).resolve().parent.parent
    pastas = [
        _nc / "assets" / "Template",
        _nc / "assets" / "templates",
    ]
    for pasta in pastas:
        for nome in ("Template_EAF.xlsx", "Template_EAF.xlsx.xlsx"):
            candidato = pasta / nome
            if candidato.is_file():
                return candidato
    return M01_TEMPLATE_EAF  # usado na mensagem de erro se não encontrar nenhum


# ESTRUTURA DA PLANILHA NO SEPARAR NC (fixo):
#   Linhas 1 a 4 = cabeçalho (nunca gravar dados aqui).
#   A partir da linha 5 = dados (sempre gravar linhas do grupo a partir da linha 5).
# ─────────────────────────────────────────────────────────────────────────────
LINHA_CABECALHO_FIM = 4   # última linha do cabeçalho (1–4)
PRIMEIRA_LINHA_DADOS = 5  # gravar dados sempre a partir da linha 5

# CONSTANTES DE COLUNAS – planilha-mãe EAF (índice 1 = col A)
# Colunas fixas (iguais em todas as versões do EAF):
#   C=código, D=data constatação, F=rodovia, I=m_ini, K=m_fim, Q=Atividade, V=nº foto
# Coluna variável detectada dinamicamente no cabeçalho:
#   "Data Reparo" → T(20) no template manual | S(19) nos exports do sistema ARTESP
# Demais colunas (para preenchimento completo pelo módulo MA):
#   G=concessionária/EAF, H=km inicial (formato 143+800), J=km final, L=sentido, O=tipo atividade, P=grupo, U=responsável
# ─────────────────────────────────────────────────────────────────────────────
COL_KM_I_M   = 9   # I – metros inicial
COL_KM_F_M   = 11  # K – metros final
COL_CODIGO   = 3   # C – código fiscalização / número da NC
COL_SEQ_FOTO = 22  # V – número da NC para foto (código da col C, ou sequencial)
COL_DATA_NC  = 20  # T – data reparo/prazo (fallback; detectado dinamicamente em executar())
COL_RODOVIA  = 6   # F – rodovia
COL_TIPO_NC  = 17  # Q – tipo/serviço NC (Atividade)
COL_DATA_CON = 4   # D – data da constatação
COL_CONCESSIONARIA = 7   # G – concessionária / EAF
COL_KM_I_FULL     = 8   # H – km inicial (formato 143+800)
COL_KM_F_FULL     = 10  # J – km final (formato 143+800)
COL_SENTIDO       = 12  # L – sentido
COL_TIPO_ATIV     = 15  # O – tipo atividade
COL_GRUPO_ATIV    = 16  # P – grupo (fiscalização)
COL_RESPONSAVEL   = 21  # U – responsável (fiscal)


def _detectar_col_data_reparo(ws, fallback: int = 20) -> int:
    """
    Detecta a coluna 'Data Reparo' lendo o cabeçalho da planilha (linhas 1-5).
    Compatível com:
      - Template manual (_Planilha Modelo nc lote 13.xls): col T(20) = 'Data Reparo'
      - Exports do sistema ARTESP: col S(19) = 'Data Reparo'
    Retorna o índice 1-based encontrado, ou `fallback` se não localizar.
    """
    for r in range(1, 6):
        for c in range(1, min(ws.max_column + 1, 30)):
            v = ws.cell(row=r, column=c).value
            if v is not None and str(v).strip().lower() == "data reparo":
                logger.debug(f"Coluna 'Data Reparo' detectada: col {c} (linha {r})")
                return c
    logger.warning(f"Coluna 'Data Reparo' nao encontrada no cabecalho — usando fallback col {fallback}")
    return fallback


def _detectar_col_data_envio(ws, fallback: int = 19) -> int:
    """
    Detecta a coluna 'Data do envio' ou 'Data envio' no cabeçalho (linhas 1-5).
    Usado no template EAF para mapear data da fiscalização (constatação) para a coluna correta.
    Retorna o índice 1-based; se não encontrar, retorna fallback.
    """
    for r in range(1, 6):
        for c in range(1, min(ws.max_column + 1, 30)):
            v = ws.cell(row=r, column=c).value
            if v is None:
                continue
            s = str(v).strip().lower()
            if "data" in s and "envio" in s:
                logger.debug(f"Coluna 'Data do envio' detectada: col {c} (linha {r})")
                return c
    return fallback


def _detectar_colunas_data_no_template(ws_template) -> tuple[int | None, int | None]:
    """
    Detecta no template EAF (cabeçalho linhas 1-5) as colunas 'Data do envio' e 'Data do reparo'.
    Retorna (col_data_envio, col_data_reparo). Se alguma não for encontrada, retorna None para essa.
    Fallback: se só 'Data Reparo' for encontrada em T(20), assume S(19) = Data do envio.
    """
    col_envio, col_reparo = None, None
    for r in range(1, 6):
        for c in range(1, min(ws_template.max_column + 1, 30)):
            v = ws_template.cell(row=r, column=c).value
            if v is None:
                continue
            s = str(v).strip().lower()
            if "data" in s and "envio" in s:
                col_envio = c
            if s == "data reparo":
                col_reparo = c
        if col_envio is not None and col_reparo is not None:
            break
    # Template pode ter só "Data Reparo" (ex.: col T); S(19) = Data do envio por convenção
    if col_reparo is not None and col_envio is None and col_reparo == 20:
        col_envio = 19
    return (col_envio, col_reparo)


def _converter_xls_para_xlsx(path_xls: Path) -> Path:
    """
    Lê um arquivo .xls (formato antigo) com xlrd e grava um .xlsx equivalente
    com openpyxl no mesmo diretório. Cabeçalho (linhas 1 a M01_LINHA_INICIO-1)
    e demais linhas são copiadas só com valores; não inventa cabeçalho genérico.
    Retorna o Path do arquivo .xlsx gerado.
    """
    path_xlsx = path_xls.with_suffix(".xlsx")
    if path_xlsx == path_xls:
        path_xlsx = path_xls.parent / (path_xls.stem + "_convertido.xlsx")

    book = xlrd.open_workbook(str(path_xls))
    sheet = book.sheet_by_index(0)

    wb = openpyxl.Workbook()
    ws = wb.active
    if sheet.name:
        ws.title = sheet.name[:31]  # limite de 31 caracteres no Excel

    # Copiar só os valores do .xls (cabeçalho e dados); não gravar nada genérico.
    for row in range(sheet.nrows):
        for col in range(sheet.ncols):
            cell = sheet.cell(row, col)
            if cell.ctype == xlrd.XL_CELL_EMPTY:
                continue
            excel_cell = ws.cell(row=row + 1, column=col + 1)
            if cell.ctype == xlrd.XL_CELL_NUMBER:
                excel_cell.value = cell.value
            elif cell.ctype == xlrd.XL_CELL_DATE:
                try:
                    dt = xlrd.xldate.xldate_as_datetime(cell.value, book.datemode)
                    excel_cell.value = dt
                except Exception:
                    excel_cell.value = cell.value
            elif cell.ctype == xlrd.XL_CELL_TEXT:
                excel_cell.value = cell.value
            elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
                excel_cell.value = bool(cell.value)
            else:
                excel_cell.value = cell.value

    wb.save(str(path_xlsx))
    logger.info(f"Arquivo .xls convertido para: {path_xlsx.name}")
    return path_xlsx


def _cell(ws, row: int, col: int):
    """Retorna valor da célula (row, col) ou string vazia."""
    v = ws.cell(row=row, column=col).value
    return v if v is not None else ""


def _limpar_str(val) -> str:
    """Normaliza valor para comparação (string limpa)."""
    return str(val).strip() if val is not None else ""


def _copiar_linha_com_estilo(ws_src, row_src: int, ws_dst, row_dst: int, max_col: int) -> None:
    """
    Copia uma linha inteira (valor + estilo) de ws_src para ws_dst.
    Mesma regra do gerador de modelo/foto: preserva font, border, fill, alignment, number_format.
    """
    for col in range(1, max_col + 1):
        src_cell = ws_src.cell(row=row_src, column=col)
        dst_cell = ws_dst.cell(row=row_dst, column=col)
        dst_cell.value = src_cell.value
        if src_cell.has_style:
            dst_cell.font = src_cell.font.copy()
            dst_cell.border = src_cell.border.copy()
            dst_cell.fill = src_cell.fill.copy()
            dst_cell.number_format = src_cell.number_format
            dst_cell.alignment = src_cell.alignment.copy()


def _copiar_alturas_linhas(ws_src, ws_dst, src_start: int, num_linhas: int, dst_start: int) -> None:
    """Copia as alturas das linhas do bloco origem para o bloco destino (como em gerar_modelo_foto)."""
    for offset in range(num_linhas):
        dim = ws_src.row_dimensions.get(src_start + offset)
        if dim is not None and dim.height is not None:
            ws_dst.row_dimensions[dst_start + offset].height = dim.height


def _replicar_merged_cells_header(ws_src, ws_dst, row_ini: int, row_fim: int) -> None:
    """
    Replica no sheet destino as células mescladas que estão inteiras no range [row_ini, row_fim] do origem.
    Destino usa as mesmas coordenadas (cabeçalho 1:row_fim).
    """
    for mc in list(ws_src.merged_cells.ranges):
        if mc.min_row >= row_ini and mc.max_row <= row_fim:
            try:
                ws_dst.merge_cells(
                    start_row=mc.min_row,
                    start_column=mc.min_col,
                    end_row=mc.max_row,
                    end_column=mc.max_col,
                )
            except Exception:
                pass


def atualizar_col_v_indice_global(arqs: list[Path], start_index: int) -> int:
    """
    Atualiza a coluna V (número da foto) nos XLS individuais com índice global
    sequencial (start_index+1, start_index+2, ...), na ordem dos arquivos e das linhas.
    Usado após a extração de fotos do PDF para que M02 encontre PDF (1).jpg, PDF (2).jpg, etc.
    Retorna o número de linhas atualizadas.
    """
    from openpyxl import load_workbook
    idx = start_index
    total = 0
    for path in arqs:
        if not path.exists():
            continue
        wb = load_workbook(str(path))
        ws = wb.active
        ultima = ws.max_row
        # Encontrar última linha com dados na col C
        for r in range(ultima, M01_LINHA_INICIO - 1, -1):
            if ws.cell(row=r, column=COL_CODIGO).value:
                ultima = r
                break
        for r in range(M01_LINHA_INICIO, ultima + 1):
            if ws.cell(row=r, column=COL_CODIGO).value is None:
                continue
            idx += 1
            ws.cell(row=r, column=COL_SEQ_FOTO).value = idx
            total += 1
        wb.save(str(path))
        wb.close()
    return total


def _nome_arquivo(rodovia_raw: str, tipo_nc: str,
                  data_constatacao, data_prazo) -> str:
    """
    Monta o nome do arquivo exportado:
    yyyymmdd - CONSTATAÇÕES NC LOTE 13 (rod - serv) - Prazo - dd-mm-aaaa.xlsx
    """
    rod_info = normalizar_rodovia_eaf(rodovia_raw, RODOVIAS)
    # VBA: rod = Left(F,6); se "SPI 10" então "SPI 102-300"; senão usa rod
    rod_6    = str(rodovia_raw).strip()[:6]
    rod_nome = RODOVIA_NOME_SEPARAR.get(rod_info["tag"], rod_6) if rod_info["tag"] != "FORA" else sanitizar_nome(str(rodovia_raw)[:10])

    # Abreviação do serviço
    serv_abrev = SERVICO_ABREV.get(tipo_nc.strip(), sanitizar_nome(tipo_nc[:30]))

    # Datas: evita 00000000/00-00-0000 no nome usando data de hoje como fallback
    dt_con  = parse_data(data_constatacao)
    dt_praz = parse_data(data_prazo)
    if not dt_con:
        dt_con = datetime.now()
    if not dt_praz:
        dt_praz = datetime.now()
    yyyymmdd = data_yyyymmdd(dt_con)
    prazo_s  = data_ddmmaaaa(dt_praz)

    nome = (
        f"{yyyymmdd} - CONSTATAÇÕES NC {M01_LOTE} "
        f"({rod_nome} - {serv_abrev}) - Prazo - {prazo_s}.xlsx"
    )
    return sanitizar_nome(nome)


def executar(arquivo_mae: Path, pasta_destino: Path | None = None,
             callback_progresso=None, sobrescrever: bool = False,
             um_arquivo_por_nc: bool = False) -> list[Path]:
    """
    Processa a planilha-mãe EAF e gera os arquivos individuais de NC.
    sobrescrever: se True, regrava arquivos que já existem (útil em testes locais).
    um_arquivo_por_nc: se True, gera um Excel por linha (uma NC por arquivo); senão agrupa por (data, rodovia, tipo).

    Parâmetros
    ----------
    arquivo_mae      : Path para o .xls/.xlsx mãe (planilha EAF completa).
    pasta_destino    : Pasta onde os XLS serão salvos (padrão: M01_EXPORTAR).
    callback_progresso : função(atual, total, msg) para atualizar GUI.
    um_arquivo_por_nc : se True, um arquivo por linha da EAF (uma NC por Excel).

    Retorna
    -------
    Lista de Path dos arquivos gerados.
    """
    pasta_destino = Path(pasta_destino) if pasta_destino else M01_EXPORTAR
    garantir_pasta(pasta_destino)

    if not arquivo_mae.is_file():
        raise ValueError(
            "O caminho informado é uma pasta, não um arquivo. "
            "Selecione o arquivo Excel da planilha EAF (ex.: L13 - NC Constatação...xls ou Planilha_EAF_mae_teste.xlsx)."
        )

    suff = arquivo_mae.suffix.lower()
    if suff == ".pdf":
        raise ValueError(
            "O arquivo selecionado é um PDF. O passo [1/6] Separando NCs exige a PLANILHA EXCEL (planilha-mãe EAF), não o PDF.\n\n"
            "Selecione o arquivo .xlsx ou .xls da planilha EAF (ex.: Lote 13 - SP 075 - 13_02_2026_Conservação de Rotina.xlsx). "
            "O PDF é usado em outra etapa (extração de imagens), não aqui."
        )

    # openpyxl só lê .xlsx/.xlsm; converter .xls quando necessário
    if suff == ".xls":
        arquivo_mae = _converter_xls_para_xlsx(arquivo_mae)
    elif suff not in (".xlsx", ".xlsm", ".xltx", ".xltm"):
        # Sem extensão ou extensão desconhecida: tentar .xls (xlrd) primeiro
        try:
            xlrd.open_workbook(str(arquivo_mae))
            arquivo_mae = _converter_xls_para_xlsx(arquivo_mae)
        except Exception:
            # Pode ser .xlsx salvo sem extensão – deixar load_workbook tentar
            pass

    # Abrir só para ler: detectar colunas, última linha e lista de (data, rodovia, tipo) por linha.
    # Não gravar no arquivo-mãe — cada saída será cópia binária do original (cabeçalho intacto).
    logger.info(f"Abrindo planilha-mãe (somente leitura): {arquivo_mae.name}")
    wb_mae = load_workbook(str(arquivo_mae), read_only=False)
    ws = wb_mae.active

    col_data_reparo = _detectar_col_data_reparo(ws, fallback=COL_DATA_NC)
    logger.info(f"Coluna 'Data Reparo': {col_data_reparo}")

    # Última linha com dado: procura de baixo para cima onde col C tem valor; se EAF tem poucas linhas, max_row já é correto
    ultima_linha = M01_LINHA_INICIO - 1
    for r in range(ws.max_row, M01_LINHA_INICIO - 1, -1):
        if ws.cell(row=r, column=COL_CODIGO).value:
            ultima_linha = r
            break

    total_linhas = ultima_linha - M01_LINHA_INICIO + 1
    logger.info(f"Linhas de dados: {total_linhas} (L{M01_LINHA_INICIO}–L{ultima_linha})")

    # Lista (data_nc, rodov_raw, tipo_nc, data_con) por linha para o loop
    # Se Tipo NC (Q) estiver vazio (ex.: EAF gerada desde PDF MA), usa código (C) ou "NC" para não pular a linha
    linhas_info = []
    for r in range(M01_LINHA_INICIO, ultima_linha + 1):
        tipo_nc = _cell(ws, r, COL_TIPO_NC)
        if not tipo_nc or not str(tipo_nc).strip():
            tipo_nc = _cell(ws, r, COL_CODIGO) or "NC"
        linhas_info.append((
            _cell(ws, r, col_data_reparo),
            _cell(ws, r, COL_RODOVIA),
            tipo_nc,
            _cell(ws, r, COL_DATA_CON),
        ))

    template_eaf = _caminho_template_eaf()
    if not template_eaf.is_file():
        raise FileNotFoundError(
            f"Template EAF não encontrado.\n"
            "Coloque Template_EAF.xlsx em nc_artesp/assets/Template/ ou nc_artesp/assets/templates/"
        )
    logger.info(f"Usando template: {template_eaf.name}")
    max_col = ws.max_column

    # Gerar arquivos: cópia do template + preenchimento com linhas do grupo da planilha-mãe
    arquivos_gerados: list[Path] = []
    processadas = set()

    for idx, r in enumerate(range(M01_LINHA_INICIO, ultima_linha + 1)):
        data_nc, rodov_raw, tipo_nc, data_con = linhas_info[idx]

        if not tipo_nc:
            continue

        nome_arq = _nome_arquivo(rodov_raw, tipo_nc, data_con, data_nc)
        if um_arquivo_por_nc:
            # Um Excel por NC: padrão do nome + código (col C); sufixo _1, _2 se repetir ou para destino único
            codigo = _cell(ws, r, COL_CODIGO)
            codigo_safe = sanitizar_nome(str(codigo).strip())[:80] if codigo and str(codigo).strip() else f"NC-{r}"
            stem, ext = Path(nome_arq).stem, Path(nome_arq).suffix
            nome_base = f"{stem} - {codigo_safe}{ext}"
            nome_arq = sanitizar_nome(nome_base)
            n = 1
            while nome_arq in processadas:
                nome_arq = sanitizar_nome(f"{stem} - {codigo_safe}_{n}{ext}")
                n += 1
            processadas.add(nome_arq)
        destino = encurtar_nome_em_pasta(pasta_destino, nome_arq)
        garantir_pasta(pasta_destino)

        if um_arquivo_por_nc:
            # Garantir destino único: se truncou e já existe, acrescenta _2, _3 até achar caminho livre
            cont = 1
            while destino.exists() and not sobrescrever:
                stem, ext = Path(nome_arq).stem, Path(nome_arq).suffix
                nome_arq = sanitizar_nome(f"{stem}_{cont}{ext}")
                processadas.add(nome_arq)
                destino = encurtar_nome_em_pasta(pasta_destino, nome_arq)
                cont += 1
        elif destino.exists() and not sobrescrever:
            logger.debug(f"Já existe, pulando: {nome_arq}")
            continue
        if not um_arquivo_por_nc and nome_arq in processadas:
            continue
        if not um_arquivo_por_nc:
            processadas.add(nome_arq)

        data_str = _limpar_str(data_nc)
        rod_str  = _limpar_str(rodov_raw)[:6]
        tipo_str = _limpar_str(tipo_nc)

        # Linhas da planilha-mãe que pertencem a este grupo (ou só esta linha se um_arquivo_por_nc)
        if um_arquivo_por_nc:
            linhas_do_grupo = [r]
        else:
            linhas_do_grupo = [
                rr for rr in range(M01_LINHA_INICIO, ultima_linha + 1)
                if (_limpar_str(_cell(ws, rr, col_data_reparo)) == data_str
                    and _limpar_str(_cell(ws, rr, COL_RODOVIA))[:6] == rod_str
                    and _limpar_str(_cell(ws, rr, COL_TIPO_NC)) == tipo_str)
            ]

        logger.info(f"Gerando: {nome_arq}")

        # 1) Usar o template EAF: cópia binária (preserva cabeçalho, formatação, colunas)
        shutil.copy2(template_eaf, destino)

        wb_copia = load_workbook(str(destino))
        ws_copia = wb_copia.active

        # Detectar no template as colunas "Data do envio" e "Data do reparo" para mapear corretamente
        col_envio_tpl, col_reparo_tpl = _detectar_colunas_data_no_template(ws_copia)
        if col_envio_tpl is not None or col_reparo_tpl is not None:
            logger.debug(f"Template: col Data do envio={col_envio_tpl}, col Data do reparo={col_reparo_tpl}")

        # 2) Remover linhas de dados do template (se houver), deixando só o cabeçalho 1–4
        while ws_copia.max_row >= PRIMEIRA_LINHA_DADOS:
            ws_copia.delete_rows(ws_copia.max_row, 1)

        # 3) Preencher a partir da linha 5 com as linhas do grupo da planilha-mãe (com estilo + padronização I, K, V)
        for seq, row_origem in enumerate(linhas_do_grupo, start=1):
            row_dest = PRIMEIRA_LINHA_DADOS + seq - 1
            _copiar_linha_com_estilo(ws, row_origem, ws_copia, row_dest, max_col)
            # Corrigir S e T pelo template: Data do envio = data da fiscalização (col D); Data do reparo = col mãe ou envio + 10 dias
            val_envio = ws.cell(row=row_origem, column=COL_DATA_CON).value
            val_reparo = ws.cell(row=row_origem, column=col_data_reparo).value
            if col_envio_tpl is not None:
                ws_copia.cell(row=row_dest, column=col_envio_tpl).value = val_envio
            if col_reparo_tpl is not None:
                if val_reparo is None or not _limpar_str(val_reparo):
                    dt_envio = parse_data(val_envio)
                    if dt_envio:
                        dt_reparo = dt_envio + timedelta(days=PRAZO_DIAS_APOS_ENVIO)
                        ws_copia.cell(row=row_dest, column=col_reparo_tpl).value = data_ddmmaaaa(dt_reparo)
                    else:
                        ws_copia.cell(row=row_dest, column=col_reparo_tpl).value = val_reparo
                else:
                    ws_copia.cell(row=row_dest, column=col_reparo_tpl).value = val_reparo
            ws_copia.cell(row=row_dest, column=COL_KM_I_M).number_format = "@"
            ws_copia.cell(row=row_dest, column=COL_KM_I_M).value = pad_metros(
                ws_copia.cell(row=row_dest, column=COL_KM_I_M).value
            )
            ws_copia.cell(row=row_dest, column=COL_KM_F_M).number_format = "@"
            ws_copia.cell(row=row_dest, column=COL_KM_F_M).value = pad_metros(
                ws_copia.cell(row=row_dest, column=COL_KM_F_M).value
            )
            codigo_raw = ws_copia.cell(row=row_dest, column=COL_CODIGO).value
            try:
                codigo_int = int(float(str(codigo_raw).strip())) if codigo_raw is not None and str(codigo_raw).strip() else None
            except (ValueError, TypeError):
                codigo_int = None
            ws_copia.cell(row=row_dest, column=COL_SEQ_FOTO).value = codigo_int if codigo_int is not None else seq

        destino_xls = destino.with_suffix(".xls")
        if destino_xls.exists():
            destino_xls.unlink()
            logger.debug(f"Removido .xls antigo: {destino_xls.name}")

        wb_copia.save(str(destino))
        wb_copia.close()
        arquivos_gerados.append(destino)
        logger.info(f"  ✓ Salvo: {destino.name}")

    wb_mae.close()

    logger.info(f"Módulo 01 concluído. {len(arquivos_gerados)} arquivo(s) gerado(s).")
    if callback_progresso:
        callback_progresso(total_linhas, total_linhas, "Módulo 01 concluído.")
    return arquivos_gerados
