"""
nc_artesp/utils/helpers.py
────────────────────────────────────────────────────────────────────────────
Funções utilitárias do pipeline NC ARTESP.
"""

from __future__ import annotations

import logging
import os
import re
import shutil
from datetime import datetime
from pathlib import Path
from typing import Optional


def configurar_log(nivel: int = logging.INFO,
                   arquivo: "Path | None" = None) -> None:
    """Configura o logging raiz com formatação padrão."""
    handlers = [logging.StreamHandler()]
    if arquivo:
        try:
            handlers.append(logging.FileHandler(str(arquivo), encoding="utf-8"))
        except Exception:
            pass
    logging.basicConfig(
        level=nivel,
        format="%(asctime)s [%(levelname)s] %(name)s — %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        handlers=handlers,
        force=True,
    )


# ─────────────────────────────────────────────────────────────────────────────
# DATAS
# ─────────────────────────────────────────────────────────────────────────────

def parse_data(valor) -> Optional[datetime]:
    """Parseia datas em vários formatos; retorna datetime ou None."""
    if valor is None:
        return None
    if isinstance(valor, datetime):
        return valor
    # date → datetime
    try:
        from datetime import date
        if isinstance(valor, date):
            return datetime(valor.year, valor.month, valor.day)
    except Exception:
        pass
    # número serial Excel (float/int)
    if isinstance(valor, (int, float)):
        try:
            import xlrd
            return datetime(*xlrd.xldate_as_tuple(float(valor), 0)[:6])
        except Exception:
            pass
        try:
            from openpyxl.utils.datetime import from_excel
            return from_excel(valor)
        except Exception:
            pass
    s = str(valor).strip()
    if not s or s.lower() in ("none", "nan", ""):
        return None
    for fmt in (
        "%d/%m/%Y", "%d/%m/%y",
        "%Y-%m-%d", "%Y%m%d",
        "%d-%m-%Y", "%d-%m-%y",
        "%d/%m/%Y %H:%M", "%d/%m/%Y %H:%M:%S",
        "%Y-%m-%d %H:%M:%S",
    ):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass
    return None


def data_yyyymmdd(dt: Optional[datetime]) -> str:
    """datetime → 'YYYYMMDD'. Retorna '00000000' se None."""
    if not dt:
        return "00000000"
    return dt.strftime("%Y%m%d")


def data_ddmmaaaa(dt: Optional[datetime]) -> str:
    """datetime → 'DD/MM/AAAA'. Retorna '' se None."""
    if not dt:
        return ""
    return dt.strftime("%d/%m/%Y")


def data_br(dt: Optional[datetime]) -> str:
    """datetime → 'DD/MM/YYYY'. Alias de data_ddmmaaaa."""
    return data_ddmmaaaa(dt)


def timestamp_agora() -> str:
    """Retorna 'YYYYMMDD-HHMM'."""
    return datetime.now().strftime("%Y%m%d-%H%M")


def timestamp_completo() -> str:
    """Retorna 'YYYYMMDD - HHMMSS'."""
    return datetime.now().strftime("%Y%m%d - %H%M%S")


# ─────────────────────────────────────────────────────────────────────────────
# KM E METROS
# ─────────────────────────────────────────────────────────────────────────────

def pad_metros(valor) -> str:
    """Normaliza metros para 3 dígitos ('50' → '050', '1000' → '000')."""
    if valor is None:
        return "000"
    s = str(valor).strip()
    # Remove parte decimal se presente
    s = s.split(".")[0].split(",")[0]
    # Mantém só dígitos
    s = re.sub(r"\D", "", s)
    if not s:
        return "000"
    # Trunca para os últimos 3 dígitos (metro 1000 → 000)
    return s[-3:].zfill(3)


def km_mais_metros(km, metros) -> str:
    """'50 + 950' a partir de km=50 e metros='950'."""
    try:
        km_s = str(int(float(str(km).replace(",", "."))))
    except Exception:
        km_s = str(km)
    met_s = pad_metros(metros)
    return f"{km_s} + {met_s}"


def km_virgula_metros(km, metros) -> str:
    """'50,950'."""
    try:
        km_s = str(int(float(str(km).replace(",", "."))))
    except Exception:
        km_s = str(km)
    met_s = pad_metros(metros)
    return f"{km_s},{met_s}"


def km_formato_arquivo(km, metros=None) -> str:
    """
    Retorna KM no formato '50+950' (sem espaços) para nomes de arquivo.
    Aceita:
      - km_formato_arquivo(50, 950)         → '50+950'
      - km_formato_arquivo('50 + 950')      → '50+950'
      - km_formato_arquivo('50+950')        → '50+950'
      - km_formato_arquivo(50.950)          → '50+950'
    """
    if metros is not None:
        return km_mais_metros(km, metros).replace(" ", "")
    # Argumento único: pode ser string formatada ou float
    s = str(km).strip()
    # Se já tem '+' ou ',', limpa espaços
    if '+' in s or ',' in s:
        return s.replace(" ", "").replace(",", "+")
    # Float → converte
    try:
        v = float(s.replace(",", "."))
        km_int = int(v)
        met = round((v - km_int) * 1000)
        return f"{km_int}+{met:03d}"
    except Exception:
        return s.replace(" ", "")


def formatar_numero(n, largura_ou_decimais: int = 3) -> str:
    """
    Formata número. Quando usado no pipeline NC ARTESP, formata como zero-padded int
    (ex: formatar_numero(1, 6) → '000001').
    Com valor grande (>= 1000) ou float, formata com casas decimais.
    """
    try:
        v = float(n)
        # Se o argumento é <= 9 provavelmente é largura (padrão do pipeline)
        if largura_ou_decimais <= 9 and v == int(v):
            return str(int(v)).zfill(largura_ou_decimais)
        return f"{v:.{largura_ou_decimais}f}"
    except Exception:
        return str(n)


# ─────────────────────────────────────────────────────────────────────────────
# NOMES DE ARQUIVO E PASTAS
# ─────────────────────────────────────────────────────────────────────────────

_CHARS_INVALIDOS = re.compile(r'[\\/:*?"<>|\x00-\x1f]')


def sanitizar_nome(s: str, max_len: int = 200) -> str:
    """Remove caracteres inválidos para nome de arquivo Windows."""
    if not s or not isinstance(s, str):
        return ""
    s = _CHARS_INVALIDOS.sub("_", s)
    s = s.strip(". ")
    return s[:max_len]


def garantir_pasta(caminho) -> Path:
    """Cria o diretório se não existir. Retorna Path."""
    p = Path(caminho)
    p.mkdir(parents=True, exist_ok=True)
    return p


def caminho_dentro_limite_windows(caminho, max_len: int = 259) -> Path:
    """
    Retorna o caminho garantindo que não ultrapasse max_len chars.
    Se o caminho for longo demais, trunca o nome do arquivo preservando a extensão.
    """
    p = Path(caminho)
    if len(str(p)) <= max_len:
        return p
    ext    = p.suffix
    base   = p.stem
    margem = max_len - len(str(p.parent)) - len(ext) - 2
    margem = max(margem, 5)
    return p.parent / (base[:margem] + ext)


def encurtar_nome_em_pasta(pasta: Path, nome: str, max_path: int = 259) -> Path:
    """
    Retorna Path(pasta/nome) encurtando o nome se o caminho total exceder max_path.
    Preserva a extensão.
    """
    destino = pasta / nome
    if len(str(destino)) <= max_path:
        return destino
    ext  = Path(nome).suffix
    base = Path(nome).stem
    margem = max_path - len(str(pasta)) - len(ext) - 2  # -2 para / e margem
    if margem < 10:
        margem = 10
    nome_curto = base[:margem] + ext
    return pasta / nome_curto


def copiar_arquivo(src, dst, sobrescrever: bool = True) -> Path:
    """Copia arquivo src → dst. Cria pasta de destino se necessário."""
    src = Path(src);  dst = Path(dst)
    if not sobrescrever and dst.exists():
        return dst
    garantir_pasta(dst.parent)
    shutil.copy2(str(src), str(dst))
    return dst


def renomear_arquivo(src, dst) -> Path:
    """Renomeia/move src → dst. Se o destino já existir (reprocessamento), remove ou substitui.
    No Windows/OneDrive o destino pode persistir após unlink; usa os.replace como fallback."""
    src = Path(src)
    dst = Path(dst)
    garantir_pasta(dst.parent)
    if dst.exists():
        try:
            dst.unlink()
        except OSError:
            pass
    try:
        src.rename(dst)
    except FileExistsError:
        # WinError 183 / OneDrive: destino ainda existe; substituir explicitamente
        os.replace(str(src), str(dst))
    return dst


# ─────────────────────────────────────────────────────────────────────────────
# RODOVIAS E MAPA EAF (grupos por trecho km — Contatos EAFs)
# ─────────────────────────────────────────────────────────────────────────────

def normalizar_rodovia_para_busca(rodovia: str) -> str:
    """
    Normaliza o nome da rodovia para comparação com MAPA_EAF e RODOVIAS.
    Aceita: "SP 075", "SP075", "SP-075", "SP 127", "SPI 102-300", etc.
    Usado por: análise de PDF (atribuir grupo), e-mail, separar NC, helpers.
    """
    s = str(rodovia or "").strip().upper()
    s = re.sub(r"\s+", " ", s)
    s = s.replace("-", " ").replace("_", " ")
    s = re.sub(r"\bSP(\d)", r"SP \1", s)
    s = re.sub(r"\bSPI(\d)", r"SPI \1", s)
    return s


def obter_grupo_empresa_por_trecho(rodovia: str, km: float, mapa_eaf: list) -> tuple[int, str]:
    """
    Retorna (grupo, empresa) para uma NC com base em rodovia + km, usando MAPA_EAF.
    mapa_eaf: list[dict] com keys grupo, empresa, trechos (lista de {rodovia, km_ini, km_fim}).
    Retorna (0, "") se não houver trecho correspondente.
    """
    if not mapa_eaf or km is None:
        return 0, ""
    rod_nc = normalizar_rodovia_para_busca(rodovia)
    if not rod_nc:
        return 0, ""
    for entry in mapa_eaf:
        for trecho in entry.get("trechos", []):
            rod_t = normalizar_rodovia_para_busca(trecho.get("rodovia", ""))
            if not rod_t:
                continue
            if rod_t == rod_nc or rod_t in rod_nc or rod_nc in rod_t:
                ki = trecho.get("km_ini", 0.0)
                kf = trecho.get("km_fim", 9999.0)
                if ki <= km <= kf:
                    return entry.get("grupo", 0), entry.get("empresa", "")
    return 0, ""


def normalizar_rodovia_eaf(rodovia_raw: str, rodovias: dict) -> dict:
    """
    Normaliza o nome de rodovia da EAF buscando em `rodovias` (config.RODOVIAS).
    Retorna dict com keys 'tag', 'nome', 'sentidos', 'codigo', 'n'.
    Se não encontrar, retorna tag='FORA' com o raw como nome.
    """
    raw = str(rodovia_raw or "").strip()

    def _completar(info: dict, chave: str) -> dict:
        tag = info.get("tag", chave)
        # codigo: forma exibição (ex. SP-075); se chave é "SP 075" -> "SP-075"
        codigo = info.get("codigo") or chave.replace(" ", "-") if chave else tag
        n = info.get("n", 0)
        return {**info, "tag": tag, "codigo": codigo, "n": n}

    # Busca exata
    if raw in rodovias:
        return _completar(rodovias[raw].copy(), raw)
    # Busca por prefixo (primeiros 6 chars)
    prefixo = raw[:6]
    for chave, info in rodovias.items():
        if chave.startswith(prefixo) or prefixo.startswith(chave[:6]):
            return _completar(info.copy(), chave)
    # Busca case-insensitive
    raw_up = raw.upper()
    for chave, info in rodovias.items():
        if chave.upper() in raw_up or raw_up in chave.upper():
            return _completar(info.copy(), chave)
    return {"tag": "FORA", "nome": raw or "Desconhecida", "sentidos": [], "codigo": raw or "FORA", "n": 0}


# ─────────────────────────────────────────────────────────────────────────────
# CAMINHOS DE FOTOS (usados por gerar_modelo_foto e inserir_nc_kria)
# ─────────────────────────────────────────────────────────────────────────────

def path_foto_nc(pasta_nc, numero: "int | str") -> Path:
    """Retorna Path para 'nc (N).jpg' na pasta. N = número ou código (ex: HE.13.0111 para MA)."""
    return Path(pasta_nc) / f"nc ({numero}).jpg"


def path_foto_pdf(pasta_pdf, numero: "int | str") -> Path:
    """Retorna Path para 'PDF (N).jpg' na pasta. N = número ou código (ex: HE.13.0111 para MA)."""
    return Path(pasta_pdf) / f"PDF ({numero}).jpg"


def encontrar_foto_por_codigo_ou_numero(
    pasta: "Path",
    prefixo: str,
    codigo: str | int | None = None,
    numero: int | None = None,
) -> "Path | None":
    """
    Encontra arquivo de foto na pasta por código ou número.
    prefixo: "nc" ou "PDF" (para nc (N).jpg ou PDF (N).jpg).
    Aceita número (1, 00001) ou código alfanumérico (ex: HE.13.0111 — Meio Ambiente).
    Retorna Path do primeiro que existir, ou None.
    """
    pasta = Path(pasta)
    if not pasta.is_dir():
        return None
    prefix = f"{prefixo} ("
    suffix = ").jpg"

    def _buscar_por_codigo_str(cod: str) -> "Path | None":
        cod = (cod or "").strip()
        if not cod:
            return None
        exacto = pasta / f"{prefixo} ({cod}){suffix}"
        if exacto.is_file():
            return exacto
        try:
            for f in sorted(pasta.iterdir()):
                if not f.is_file() or not f.name.lower().endswith(".jpg"):
                    continue
                if f.name.startswith(prefix + cod + ")") or f.name.startswith(prefix + cod + "_"):
                    return f
        except OSError:
            pass
        return None

    # Código alfanumérico (ex: HE.13.0111)
    if codigo is not None and isinstance(codigo, str) and codigo.strip():
        try:
            int(float(codigo.strip()))
        except (ValueError, TypeError):
            r = _buscar_por_codigo_str(codigo)
            if r is not None:
                return r

    for valor in (codigo, numero):
        if valor is None:
            continue
        try:
            n = int(float(str(valor).strip()))
        except (ValueError, TypeError):
            continue
        # Exato: PDF (1).jpg ou PDF (00001).jpg
        for cod in (str(n), str(n).zfill(5)):
            exacto = pasta / f"{prefixo} ({cod}){suffix}"
            if exacto.is_file():
                return exacto
        # Com sufixo _1, _2: PDF (00001)_1.jpg
        try:
            for f in sorted(pasta.iterdir()):
                if not f.is_file() or not f.name.lower().endswith(".jpg"):
                    continue
                if not f.name.startswith(prefix):
                    continue
                rest = f.name[len(prefix):]
                if ")" in rest:
                    mid = rest.split(")")[0]
                    if mid == str(n) or mid == str(n).zfill(5):
                        return f
        except OSError:
            pass
    return None
