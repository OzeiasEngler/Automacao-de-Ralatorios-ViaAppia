"""
nc_artesp/config.py
────────────────────────────────────────────────────────────────────────────
Configurações do pipeline NC ARTESP.
Os caminhos de pasta são usados apenas no modo desktop; no modo web (API)
os parâmetros das funções sobrepõem estes defaults.
Valores sobrepõíveis via variáveis de ambiente ARTESP_*.
"""

from __future__ import annotations

import os
import re
from pathlib import Path

# ─── Helpers de leitura de env ──────────────────────────────────────────────

def _env_str(key: str, default: str) -> str:
    return (os.environ.get(key) or "").strip() or default

def _env_int(key: str, default: int) -> int:
    try:
        v = (os.environ.get(key) or "").strip()
        return int(v) if v else default
    except ValueError:
        return default

# ─── Raiz do projeto (fallback local Windows) ────────────────────────────────
_BASE = Path(_env_str("ARTESP_NC_BASE", r"C:\AUTOMAÇÃO_MACROS\Macros Kcor Ellen\artesp_nc_v2.0"))

# ═══════════════════════════════════════════════════════════════════════════
# MÓDULO 01 — SEPARAR NC
# ═══════════════════════════════════════════════════════════════════════════

M01_EXPORTAR           = _BASE / "Exportar"
M01_ARQUIVOS_ANTERIORES = _BASE / "Arquivos Anteriores"   # destino após processamento completo
M01_LINHA_INICIO        = _env_int("ARTESP_M01_LINHA_INICIO", 5)   # 5 = igual à macro (cabeçalho nas linhas 1–4)
M01_LOTE                = _env_str("ARTESP_LOTE", "LOTE 13")
# Concessionária (col G da EAF): por lote ou ARTESP_CONCESSIONARIA_NOME; não confundir com "empresa" (nome EAF/grupo, col U)
_LOTE_CONCESSIONARIA    = {"13": "Rodovias das Colinas", "21": "Rodovias do Tietê", "26": "SP Serra"}
_env_concessionaria    = (os.environ.get("ARTESP_CONCESSIONARIA_NOME") or "").strip()
_lote_num              = re.search(r"\d+", M01_LOTE)
CONCESSIONARIA_NOME     = _env_concessionaria or _LOTE_CONCESSIONARIA.get((_lote_num.group(0) if _lote_num else ""), "")
# Template EAF: somente em nc_artesp/assets/templates/ (igual às demais macros)
_template_eaf_env = _env_str("ARTESP_TEMPLATE_EAF", "")
_nc_root = Path(__file__).resolve().parent  # nc_artesp/
M01_TEMPLATE_EAF        = Path(_template_eaf_env) if _template_eaf_env else _nc_root / "assets" / "templates" / "Template_EAF.xlsx"
# Data do reparo = data do envio + N dias quando não informada
PRAZO_DIAS_APOS_ENVIO   = _env_int("ARTESP_PRAZO_DIAS_APOS_ENVIO", 10)

# ═══════════════════════════════════════════════════════════════════════════
# MÓDULO 02 — GERAR MODELO FOTO (macro Art_022)
# Saída Kria:  Arquivos/Arquivo Foto - Conserva/  |  Nome: yyyymmdd-hhmm - {Artesp}.xlsx
# Saída Resposta: _Respostas/_Relatório EAF - NC/Pendentes/  |  Nome: yyyymmdd - hhmmss - {rodovia} - dd-mm-yyyy - {nc}.xlsx
# ═══════════════════════════════════════════════════════════════════════════

M02_FOTOS_NC     = _BASE / "Imagens Provisórias"
M02_FOTOS_PDF    = _BASE / "Imagens Provisórias - PDF"
M02_SALVAR_FOTO  = _BASE / "Arquivos" / "Arquivo Foto - Conserva"
# Modelos M02: somente em nc_artesp/assets/templates/ (igual às macros)
M02_MODELO_KRIA  = _nc_root / "assets" / "templates" / "Modelo Abertura Evento Kria Conserva Rotina.xlsx"
M02_MODELO_RESP  = _nc_root / "assets" / "templates" / "Modelo.xlsx"
M02_PENDENTES    = _BASE / "_Respostas" / "_Relatório EAF - NC" / "Pendentes"
# PDF de origem para step 1.5 (opcional — pode ser None)
M02_PDF_ARQUIVO  = _env_str("ARTESP_M02_PDF_ARQUIVO", "")   # caminho fixo de PDF (single mode)
M02_PDF_ORIGEM   = _env_str("ARTESP_M02_PDF_ORIGEM",  "")

# Dimensões de referência das fotos (pixels). O tamanho efetivo é o do merged range no template:
# a imagem é redimensionada para preencher exatamente o merge (nem mais nem menos).
# nc (N).jpg: 800×500 px, resolução 222 DPI horizontal e 319 DPI vertical.
M02_FOTO_W     = _env_int("ARTESP_M02_FOTO_W",     800)   # nc (N).jpg   largura
M02_FOTO_H     = _env_int("ARTESP_M02_FOTO_H",     500)   # nc (N).jpg   altura
M02_FOTO_DPI_X = _env_int("ARTESP_M02_FOTO_DPI_X", 222)   # nc (N).jpg   resolução horizontal (DPI)
M02_FOTO_DPI_Y = _env_int("ARTESP_M02_FOTO_DPI_Y", 319)   # nc (N).jpg   resolução vertical (DPI)
M02_FOTO_PDF_W = _env_int("ARTESP_M02_FOTO_PDF_W", 480)   # PDF (N).jpg referência
M02_FOTO_PDF_H = _env_int("ARTESP_M02_FOTO_PDF_H", 202)   # PDF (N).jpg referência

# ═══════════════════════════════════════════════════════════════════════════
# MÓDULO 03 — INSERIR NC CONSERVAÇÃO (macro Art_03)
# Saída: Arquivos/Conservação/  |  Nome: yyyymmdd-hhmm - {cco}.xlsx
# ═══════════════════════════════════════════════════════════════════════════

M03_ENTRADA      = _BASE / "Arquivos" / "Arquivo Foto - Conserva"
M03_IMAGENS      = _BASE / "Imagens" / "Conservação"
# Modelo Kcor-Kria: somente em nc_artesp/assets/templates/
M03_MODELO_KCOR  = _nc_root / "assets" / "templates" / "_Planilha Modelo Kcor-Kria.XLSX"
M03_SAIDA        = _BASE / "Arquivos" / "Conservação"
M03_LINHA_INICIO = _env_int("ARTESP_M03_LINHA_INICIO", 9)   # âncora y (linha do km)
M03_BLOCO        = _env_int("ARTESP_M03_BLOCO", 5)           # linhas por bloco NC

# ═══════════════════════════════════════════════════════════════════════════
# MÓDULO 04 — JUNTAR ARQUIVOS (macro Art_04)
# Saída: Arquivos/Conservação/Acumulado/  |  Nome: yyyymmdd - hhmmss - Eventos Acumulado Artesp para Exportar Kria.xlsx
# Template padrão (macro): Acumulado/Padrão/Eventos Acumulado Artesp para Exportar Kria.xlsx
# ═══════════════════════════════════════════════════════════════════════════
# O que alimenta o acumulado: os .xlsx em Arquivos/Conservação (saída do M03).

M04_ENTRADA    = _BASE / "Arquivos" / "Conservação"   # pasta com os Kcor-Kria individuais (M03)
M04_ACUMULADO  = _BASE / "Acumulado" / "Kcor_Acumulado.xlsx"  # arquivo acumulado (rede)
M04_SAIDA      = _BASE / "Arquivos" / "Conservação" / "Acumulado"  # igual à macro
M04_NOME_SAIDA = "Eventos Acumulado Artesp para Exportar Kria.xlsx"  # nome exato do relatório (macro)
M04_MODELO_ACUMULADO = _nc_root / "assets" / "templates" / "Eventos Acumulado Artesp para Exportar Kria.xlsx"  # template padrão

CABECALHO_KCOR_KRIA = [
    "NumItem", "Origem", "Motivo", "Classificação", "Tipo",
    "Rodovia", "KMi", "KMf", "Sentido", "Local",
    "Gestor", "Executor", "Data Solicitação", "Data Suspensão",
    "DtInicio_Prog", "DtFim_Prog", "DtInicio_Exec", "DtFim_Exec",
    "Prazo", "ObsGestor", "Observações", "Diretório", "Arquivos",
    "Indicador", "Unidade",
]
NUM_COLUNAS_KCOR_KRIA = 25  # A–Y

# ═══════════════════════════════════════════════════════════════════════════
# MÓDULO 05 — INSERIR NÚMERO KRIA
# ═══════════════════════════════════════════════════════════════════════════

M05_COL_NUMERO = _env_str("ARTESP_M05_COL_NUMERO", "Y")
M05_SUFIXO     = _env_str("ARTESP_M05_SUFIXO",     "26")

# ═══════════════════════════════════════════════════════════════════════════
# MÓDULO 06 — EXPORTAR CALENDÁRIO
# ═══════════════════════════════════════════════════════════════════════════

M06_PASTA_OUTLOOK = _BASE / "Calendario"

# ═══════════════════════════════════════════════════════════════════════════
# MÓDULO 07 — INSERIR NC MEIO AMBIENTE (macro Kria2_Inserir_NC_MA_Salvar_Img)
# Saída: Arquivos/Meio Ambiente/  |  Nome: yyyymmdd-hhmm - {cco}.xlsx
# ═══════════════════════════════════════════════════════════════════════════

M07_ENTRADA      = _BASE / "Arquivos" / "Arquivo Foto - MA"
M07_IMAGENS      = _BASE / "Imagens" / "Meio Ambiente"
M07_SAIDA        = _BASE / "Arquivos" / "Meio Ambiente"
M07_MODELO_KCOR  = _nc_root / "assets" / "templates" / "_Planilha Modelo Kcor-Kria.XLSX"  # mesmo modelo do M03

# ═══════════════════════════════════════════════════════════════════════════
# MÓDULOS 1–4 EXCLUSIVOS MEIO AMBIENTE (equivalentes M01, M02, M03, M04 para MA)
# M01 MA: EAF desde PDF + Separar NC → Exportar MA
# M02 MA: Gerar Modelo Foto (Kria + Resposta) desde Exportar MA
# M03 MA: = M07 (Inserir NC / Kcor-Kria Meio Ambiente)
# M04 MA: Juntar Kcor-Kria MA → Acumulado MA
# ═══════════════════════════════════════════════════════════════════════════

M01_MA_EAF          = _BASE / "Arquivos" / "Meio Ambiente" / "EAF MA"           # planilha-mãe EAF MA
M01_MA_EXPORTAR     = _BASE / "Arquivos" / "Meio Ambiente" / "Exportar MA"       # saída Separar NC (EAF individuais)
M02_MA_KRIA         = _BASE / "Arquivos" / "Meio Ambiente" / "Arquivo Foto MA"  # saída Kria M02
M02_MA_PENDENTES    = _BASE / "Arquivos" / "Meio Ambiente" / "Pendentes MA"      # saída Resposta M02
M04_MA_ENTRADA     = _BASE / "Arquivos" / "Meio Ambiente"                        # Kcor-Kria individuais (saída M03 MA)
M04_MA_ACUMULADO   = _BASE / "Acumulado" / "Kcor_Acumulado_MA.xlsx"             # acumulado MA (rede)
M04_MA_SAIDA       = _BASE / "Arquivos" / "Meio Ambiente" / "Acumulado"          # pasta do relatório acumulado MA
M04_MA_NOME_SAIDA  = "Eventos Acumulado Artesp MA - Exportar Kria.xlsx"         # nome do relatório acumulado MA

# ═══════════════════════════════════════════════════════════════════════════
# MÓDULO 08 — SALVAR IMAGEM
# ═══════════════════════════════════════════════════════════════════════════

M08_IMAGENS_SRC   = _BASE / "Imagens" / "Conservação"
M08_DESTINO       = Path(r"D:\Apontamentos NC Artesp - Imagens Classificadas")
M08_EXPORTAR      = M08_DESTINO / "_Exportar"
M08_TIPOS_EXPORTAR = ("Pav. - Depressao", "Pav. - Pano de Rolamento")
M08_TIPO_NOME_PASTA: dict[str, str] = {}   # preenchido abaixo

# ═══════════════════════════════════════════════════════════════════════════
# E-MAIL
# ═══════════════════════════════════════════════════════════════════════════

M02_FOTOS_PDF_NC  = M02_FOTOS_PDF   # alias usado por nc_criar_email
NC_EMAIL_CC: list[str] = []         # destinatários em cópia (configurar por env se necessário)

# ═══════════════════════════════════════════════════════════════════════════
# RODOVIAS DO LOTE
# Estrutura: { "prefixo_sp": { "tag": str, "nome": str, "sentidos": list } }
# ═══════════════════════════════════════════════════════════════════════════

RODOVIAS: dict[str, dict] = {
    "SP 075": {"tag": "SP075",       "nome": "Rod. Senador José Ermírio de Moraes",  "sentidos": ["Norte", "Sul"]},
    "SP 127": {"tag": "SP127",       "nome": "Rod. João Mellão",                     "sentidos": ["Norte", "Sul"]},
    "SP 280": {"tag": "SP280",       "nome": "Rod. Castello Branco",                 "sentidos": ["Leste", "Oeste"]},
    "SP 300": {"tag": "SP300",       "nome": "Rod. Marechal Rondon",                 "sentidos": ["Leste", "Oeste"]},
    "SP 102": {"tag": "SPI102/300",  "nome": "Interligação SP-102/SP-300",           "sentidos": ["Norte", "Sul"]},
    "CP 147": {"tag": "FORA",        "nome": "Rod. Fora do Lote",                    "sentidos": []},
    "CP 308": {"tag": "FORA",        "nome": "Rod. Fora do Lote",                    "sentidos": []},
}

# Mapa tag → nome curto para nome de arquivo (M01)
RODOVIA_NOME_SEPARAR: dict[str, str] = {
    "SP075":      "SP 075",
    "SP127":      "SP 127",
    "SP280":      "SP 280",
    "SP300":      "SP 300",
    "SPI102/300": "SPI 102-300",
    "FORA":       "Fora",
}

# ═══════════════════════════════════════════════════════════════════════════
# TIPOS DE SERVIÇO — abreviações para nome de arquivo
# ═══════════════════════════════════════════════════════════════════════════

SERVICO_ABREV: dict[str, str] = {
    # Pavimentação
    "Depressão ou recalque de pequena extensão":          "Pav - Depressao",
    "Afundamento de trilha de roda":                      "Pav - Afundamento",
    "Remendo profundo ou superficial":                    "Pav - Remendo",
    "Desgaste ou desagregação do revestimento":           "Pav - Desgaste",
    "Buraco ou panela":                                   "Pav - Panela",
    "Trincas":                                            "Pav - Trincas",
    "Irregularidade transversal":                         "Pav - Irregularidade",
    "Pavimentação/ Passeio/ Alambrado":                   "Pred - Pav Passeio Alam",
    # Defesas e Sinalização
    "Defensa metálica (manutenção ou substituição)":      "Reparo Defensa",
    "Sinalização vertical (implantação ou substituição)": "Sinal Vertical",
    "Sinalização horizontal":                             "Sinal Horizontal",
    "Dispositivo de segurança":                           "Disp Seguranca",
    # Drenagem
    "Desobstrução de elemento de drenagem":               "Desobstrucao Drenagem",
    "Hidráulica/ Esgoto/ Drenagem":                       "Hidr Esg Dren",
    # Vegetação e Limpeza
    "Roçada":                                             "Rocada",
    "Capina":                                             "Capina",
    "Remoção de lixo e entulho da faixa de domínio":      "Limpeza FD",
    # Erosão e Taludes
    "Recomposição de erosão em corte / aterro":           "Eros Corte Aterro",
    "Escorregamento de taludes":                          "Taludes",
    "Deslizamento":                                       "Deslizamento",
    # Estruturas e OAE
    "Recuperação de OAE":                                 "Recup OAE",
    "Manutenção de OAE":                                  "Manut OAE",
    "Dispositivo com OAE":                                "Disp OAE",
    # Iluminação e Obras
    "Iluminação":                                         "Iluminacao",
    "Obras de Arte Especiais":                            "OAE",
    "Duplicação":                                         "Duplicacao",
    "Faixa Adicional":                                    "Faixa Adicional",
}

# ═══════════════════════════════════════════════════════════════════════════
# MAPA SERVIÇO → CLASSIFICAÇÃO KCOR
# ═══════════════════════════════════════════════════════════════════════════

# SERVICO_NC: serviço → (tipo_nc, classificação, executor)
# Usado pelo inserir_nc_kria.py (M03/M07)
SERVICO_NC: dict[str, tuple] = {
    k: ("Conservação Rotina", "Conservação Rotina", "Soluciona - Conserva")
    for k in SERVICO_ABREV
}
SERVICO_NC.update({
    "Recuperação de OAE": ("Obras",              "Obras",  "Soluciona - Obras"),
    "Manutenção de OAE":  ("Obras",              "Obras",  "Soluciona - Obras"),
    "Duplicação":         ("Obras",              "Obras",  "Soluciona - Obras"),
    "Faixa Adicional":    ("Obras",              "Obras",  "Soluciona - Obras"),
})

# ═══════════════════════════════════════════════════════════════════════════
# TIPO → PASTA DE DESTINO (M08)
# ═══════════════════════════════════════════════════════════════════════════

for _tipo, _abrev in SERVICO_ABREV.items():
    M08_TIPO_NOME_PASTA[_abrev] = _abrev

# ═══════════════════════════════════════════════════════════════════════════
# MAPA EAF — Grupos de fiscalização por trecho km (equipes por rodovia + km ini/fim)
#
# grupo   : número do grupo (coluna P da EAF)
# empresa : nome EAF (NEP, Autoroutes, EBP 22)
# trechos : lista de { rodovia, km_ini, km_fim } — cada equipe é responsável por esses trechos
# email   : (opcional) e-mail do responsável pelo apontamento desse grupo — preencher conforme
#           a imagem/planilha de contatos (e-mail de cada responsável e seu grupo EAF)
#
# Lote 13 — Colinas — Contatos EAFs (Grupo 1 NEP, 2 Autoroutes, 11 EBP 22)
# ═══════════════════════════════════════════════════════════════════════════

MAPA_EAF: list[dict] = [
    {
        "grupo": 1,
        "empresa": "NEP",
        "email": "",  # preencher com o e-mail do responsável do grupo 1 (ex.: responsavel.nep@empresa.com)
        "trechos": [
            {"rodovia": "SP 075", "km_ini": 43.000, "km_fim": 77.600},
        ],
    },
    {
        "grupo": 2,
        "empresa": "Autoroutes",
        "email": "",  # preencher com o e-mail do responsável do grupo 2
        "trechos": [
            {"rodovia": "SP 075",      "km_ini": 15.000, "km_fim":  43.000},
            {"rodovia": "SP 127",      "km_ini": 60.000, "km_fim": 105.900},
            {"rodovia": "SP 280",      "km_ini": 79.380, "km_fim": 129.600},
            {"rodovia": "SP 300",      "km_ini": 64.600, "km_fim": 103.000},
            {"rodovia": "SP 300",      "km_ini": 108.900, "km_fim": 158.650},
            {"rodovia": "SPI 102-300", "km_ini":  0.000, "km_fim":   7.900},
        ],
    },
    {
        "grupo": 11,
        "empresa": "EBP 22",
        "email": "",  # preencher com o e-mail do responsável do grupo 11
        "trechos": [
            {"rodovia": "SP 127", "km_ini":  0.000, "km_fim": 32.026},
            {"rodovia": "SP 127", "km_ini": 39.900, "km_fim": 60.000},
        ],
    },
]
