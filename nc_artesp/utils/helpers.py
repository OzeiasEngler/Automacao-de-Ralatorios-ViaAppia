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


# DATAS

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


# KM E METROS

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


# NOMES DE ARQUIVO E PASTAS

_CHARS_INVALIDOS = re.compile(r'[\\/:*?"<>|\x00-\x1f]')


def sanitizar_nome(s: str, max_len: int = 200) -> str:
    """Remove caracteres inválidos para nome de arquivo Windows."""
    if not s or not isinstance(s, str):
        return ""
    s = _CHARS_INVALIDOS.sub("_", s)
    s = s.strip(". ")
    return s[:max_len]


def truncar_nome_preservando_sufixo_prazo_m01(nome: str, max_chars: int) -> str:
    """
    Encurta o nome do ficheiro (com extensão) para no máximo ``max_chars`` caracteres,
    preservando o sufixo Art_011 / M01 `` - Prazo - dd-mm-aaaa`` antes da extensão quando existir.
    Usado no ZIP da API (componente limitado a N chars) e noutros truncamentos agressivos.
    """
    nome = (nome or "").strip()
    if not nome or len(nome) <= max_chars:
        return nome
    ext = Path(nome).suffix
    stem = Path(nome).stem
    room = max(8, max_chars - len(ext))
    if len(stem) <= room:
        return nome
    tail = ""
    m = re.search(r"( - Prazo - \d{1,2}-\d{1,2}-\d{4})$", stem)
    if m:
        tail = m.group(1)
    else:
        m = re.search(r"( - Prazo - .+)$", stem)
        if m:
            tail = m.group(1)
    if tail and len(tail) <= room:
        head_budget = room - len(tail)
        if head_budget > 0:
            head = stem[:head_budget].rstrip(" -")
            if not head:
                head = stem[:head_budget]
        else:
            head = ""
        return head + tail + ext
    if tail and len(tail) + len(ext) <= max_chars:
        return tail.strip() + ext
    return stem[:room].rstrip(" -.") + ext


def str_caminho_io_windows(caminho) -> str:
    """
    Caminho absoluto para ``open()``, ``shutil``, ``ZipFile.extractall``, openpyxl, etc. no Windows.

    Prefixa **sempre** ``\\\\?\\`` (ou ``\\\\?\\UNC\\`` em partilhas de rede) no caminho absoluto
    resolvido, permitindo até ~32 767 caracteres sem «LongPathsEnabled» e sem falhas perto do
    limite clássico de 260. Em outros SO devolve ``str(Path.resolve())``.
    """
    if os.name != "nt":
        p = Path(caminho)
        try:
            return str(p.resolve(strict=False))
        except (OSError, RuntimeError):
            return str(p)
    p = Path(caminho)
    try:
        abs_s = str(p.resolve(strict=False))
    except (OSError, RuntimeError):
        abs_s = str(p if p.is_absolute() else Path.cwd() / p)
    abs_s = os.path.normpath(abs_s)
    if abs_s.startswith("\\\\?\\"):
        return abs_s
    if abs_s.startswith("\\\\"):
        return "\\\\?\\UNC\\" + abs_s[2:].lstrip("\\")
    return "\\\\?\\" + abs_s


def str_caminho_outlook_mapi(caminho) -> str:
    """
    Caminho para ``Outlook.Application`` / ``Attachments.Add`` e outras APIs MAPI/COM.

    O Outlook costuma **falhar** com o prefixo ``\\\\?\\``; usa-se caminho clássico (sem prefixo)
    quando o comprimento absoluto ≤ 259. Acima disso volta a ``str_caminho_io_windows`` (pode ainda
    falhar no COM — nesse caso copiar o ficheiro para pasta curta).
    """
    if os.name != "nt":
        p = Path(caminho)
        try:
            return str(p.resolve(strict=False))
        except (OSError, RuntimeError):
            return str(p)
    p = Path(caminho)
    try:
        s = str(p.resolve(strict=False))
    except (OSError, RuntimeError):
        s = str(p if p.is_absolute() else Path.cwd() / p)
    s = os.path.normpath(s)
    if s.startswith("\\\\?\\"):
        return s
    if len(s) <= 259:
        return s
    return str_caminho_io_windows(p)


def extrair_zipfile_para_pasta(zf, destino) -> None:
    """
    ``ZipFile.extractall`` com pasta de destino criada via caminho estendido no Windows
    (``\\\\?\\``), para extrações profundas no servidor/pipeline.
    """
    import zipfile as _zipfile

    if not isinstance(zf, _zipfile.ZipFile):
        raise TypeError("extrair_zipfile_para_pasta espera zipfile.ZipFile")
    p = Path(destino)
    if os.name != "nt":
        p.mkdir(parents=True, exist_ok=True)
        zf.extractall(str(p))
        return
    dest_s = str_caminho_io_windows(p)
    os.makedirs(dest_s, exist_ok=True)
    zf.extractall(dest_s)


def garantir_pasta(caminho) -> Path:
    """Cria o diretório se não existir. Retorna Path. No Windows usa caminho longo se preciso."""
    p = Path(caminho)
    if os.name == "nt":
        os.makedirs(str_caminho_io_windows(p), exist_ok=True)
        return p
    p.mkdir(parents=True, exist_ok=True)
    return p


def escrever_bytes_caminho(caminho, data: bytes) -> Path:
    """Grava bytes no ficheiro; cria pastas e usa caminho longo no Windows quando necessário."""
    p = Path(caminho)
    garantir_pasta(p.parent)
    with open(str_caminho_io_windows(p), "wb") as f:
        f.write(data)
    return p


def caminho_dentro_limite_windows(caminho, max_len: int = 259) -> Path:
    """
    Retorna o caminho garantindo que não ultrapasse max_len chars (nome do ficheiro truncado).
    Se a pasta sozinha já exceder o limite, devolve o caminho original (use ``str_caminho_io_windows`` na escrita).
    """
    p = Path(caminho)
    parent_s = str(p.parent)
    ext = p.suffix
    base = p.stem
    sep = 1

    def _total(stem: str) -> int:
        return len(parent_s) + sep + len(stem) + len(ext)

    if _total(base) <= max_len:
        return p
    room = max_len - len(parent_s) - sep - len(ext)
    if room < 1:
        return p
    return p.parent / (base[:room] + ext)


def encurtar_nome_em_pasta(pasta: Path, nome: str, max_path: int = 259) -> Path:
    """
    Retorna Path(pasta/nome) encurtando o nome se o caminho total exceder max_path.
    Preserva o sufixo M01 « - Prazo - dd-mm-aaaa» antes de .xlsx (ver ``truncar_nome_preservando_sufixo_prazo_m01``).
    """
    destino = pasta / nome
    if len(str(destino)) <= max_path:
        return destino
    pasta_s = str(pasta)
    max_nome = max_path - len(pasta_s) - 1
    if max_nome < 12:
        max_nome = 12
    nome_curto = truncar_nome_preservando_sufixo_prazo_m01(nome, max_nome)
    return pasta / nome_curto


def copiar_arquivo(src, dst, sobrescrever: bool = True) -> Path:
    """Copia arquivo src → dst. Cria pasta de destino se necessário."""
    src = Path(src);  dst = Path(dst)
    if not sobrescrever and dst.exists():
        return dst
    garantir_pasta(dst.parent)
    shutil.copy2(str_caminho_io_windows(src), str_caminho_io_windows(dst))
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
        os.replace(str_caminho_io_windows(src), str_caminho_io_windows(dst))
    return dst


# RODOVIAS E MAPA EAF (grupos por trecho km — Contatos EAFs)

def normalizar_rodovia_para_busca(rodovia: str) -> str:
    """
    Normaliza o nome da rodovia para comparação com MAPA_EAF e RODOVIAS.
    Aceita: "SP 075", "SP075", "SP-075", "SP 127", "SPI 102-300", "SPI102/300", etc.
    SPI102/300 pertence à Autoroutes (trecho SPI 102-300 no MAPA_EAF).
    """
    s = str(rodovia or "").strip().upper()
    s = re.sub(r"\s+", " ", s)
    s = s.replace("-", " ").replace("_", " ").replace("/", " ")
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
    # Tolerancia mínima para contornar erros de arredondamento em floats
    # (ex.: 43.000 pode chegar como 42.999999).
    # Em PDFs, o km frequentemente vem com arredondamento/representação float.
    # Usar eps maior para não "perder" a borda exata (ex.: 43.000 lido como 42.999).
    eps = 1e-3

    def _rodovias_equivalentes(rod_trecho: str, rod_nc: str) -> bool:
        """
        Considera equivalentes rodovias como:
        - "SP 075" == "SP 75" (zeros à esquerda no número da SP)
        - "SPI 102-300" / "SPI 102 300" / "SPI 102300" (mesma rodovia)
        """
        a = normalizar_rodovia_para_busca(rod_trecho or "")
        b = normalizar_rodovia_para_busca(rod_nc or "")
        if not a or not b:
            return False
        if a == b:
            return True
        m1 = re.match(r"^(SP)\s*(\d+)$", a)
        m2 = re.match(r"^(SP)\s*(\d+)$", b)
        if m1 and m2 and m1.group(1) == m2.group(1):
            try:
                return int(m1.group(2)) == int(m2.group(2))
            except Exception:
                return False
        # SPI: comparar por sequência de dígitos (ex.: "SPI 102 300" e "SPI 102300")
        if a.startswith("SPI") and b.startswith("SPI"):
            dig_a = "".join(re.findall(r"\d+", a))
            dig_b = "".join(re.findall(r"\d+", b))
            return dig_a == dig_b
        if a.startswith("MG") and b.startswith("MG"):
            dig_a = "".join(re.findall(r"\d+", a))
            dig_b = "".join(re.findall(r"\d+", b))
            return bool(dig_a and dig_b and dig_a == dig_b)
        if a.startswith("BR") and b.startswith("BR"):
            dig_a = "".join(re.findall(r"\d+", a))
            dig_b = "".join(re.findall(r"\d+", b))
            return bool(dig_a and dig_b and dig_a == dig_b)
        return False
    # Pode haver múltiplas EAFs na mesma rodovia (e até trechos próximos).
    # Então, em vez de retornar a "primeira" correspondência, coletamos todas
    # as que contêm o km e escolhemos a mais específica.
    candidatos: list[tuple[int, str, float]] = []  # (grupo, empresa, km_ini do trecho)

    for entry in mapa_eaf:
        for trecho in entry.get("trechos", []):
            rod_t = normalizar_rodovia_para_busca(trecho.get("rodovia", ""))
            if not rod_t:
                continue
            # Rodovia deve casar por equivalência (ex.: SP 075 vs SP 75).
            if _rodovias_equivalentes(rod_t, rod_nc):
                ki = trecho.get("km_ini", 0.0)
                kf = trecho.get("km_fim", 9999.0)
                if (ki - eps) <= km <= (kf + eps):
                    candidatos.append((entry.get("grupo", 0), entry.get("empresa", ""), float(ki)))

    if not candidatos:
        return 0, ""

    # Mais específica = maior km_ini (trecho mais "tarde" na mesma rodovia).
    candidatos.sort(key=lambda x: x[2], reverse=True)
    return candidatos[0][0], candidatos[0][1]


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


# CAMINHOS DE FOTOS (usados por gerar_modelo_foto e inserir_nc_kria)

def path_foto_nc(pasta_nc, numero: "int | str") -> Path:
    """Retorna Path para 'nc (N).jpg' na pasta. N = número ou código (ex: HE.13.0111 para MA)."""
    return Path(pasta_nc) / f"nc ({numero}).jpg"


def path_foto_pdf(pasta_pdf, numero: "int | str") -> Path:
    """Retorna Path para 'PDF (N).jpg' (subpasta PDF/ se existir)."""
    p = Path(pasta_pdf)
    sub = p / "PDF"
    if sub.is_dir():
        return sub / f"PDF ({numero}).jpg"
    return p / f"PDF ({numero}).jpg"


def _pastas_busca_foto_extracao(pasta: "Path", prefixo: str) -> list:
    """Pastas candidatas para fotos extraídas (raiz, subpastas e ZIP com pasta-base)."""
    pasta = Path(pasta)
    sub = "PDF" if (prefixo or "").strip().upper() == "PDF" else "nc"
    out: list[Path] = []

    def _add(p: Path) -> None:
        if p.is_dir() and p not in out:
            out.append(p)

    if pasta.is_dir():
        # Estrutura direta (legado e alguns fluxos): raiz + raiz/sub
        _add(pasta / sub)
        _add(pasta)

        # Estrutura comum do ZIP web: pasta/lote_.../{arquivos} e/ou pasta/lote_.../sub
        try:
            for child in sorted(pasta.iterdir()):
                if not child.is_dir():
                    continue
                _add(child / sub)
                _add(child)
        except OSError:
            pass

    return out or [pasta]


_FOTO_INDEX_CACHE: dict[tuple[str, str], tuple[dict[str, Path], dict[str, Path]]] = {}
_FOTO_INDEX_RECURSIVO_CACHE: dict[tuple[str, str], tuple[dict[str, Path], dict[str, Path]]] = {}


def limpar_cache_indices_foto() -> None:
    """Limpa cache global de índices de imagens (evita paths stale após purge/extract)."""
    _FOTO_INDEX_CACHE.clear()
    _FOTO_INDEX_RECURSIVO_CACHE.clear()


def _indexar_fotos_base(base: Path, prefixo: str) -> tuple[dict[str, Path], dict[str, Path]]:
    """
    Indexa 1x os JPG de uma pasta:
      - exato: nome lower -> Path
      - mid: valor dentro de '(...)' -> Path (case-insensitive)
    """
    key = (str(base), (prefixo or "").strip().lower())
    cached = _FOTO_INDEX_CACHE.get(key)
    if cached is not None:
        return cached

    exato: dict[str, Path] = {}
    mid: dict[str, Path] = {}
    pref_l = (prefixo or "").strip().lower()
    start_l = f"{pref_l} ("

    try:
        for f in base.iterdir():
            if not f.is_file():
                continue
            name = f.name
            low = name.lower()
            if not low.endswith(".jpg"):
                continue
            exato[low] = f
            if not low.startswith(start_l):
                continue
            rest = name[len(prefixo) + 2 :]  # após "PREFIXO ("
            if ")" not in rest:
                continue
            mid_raw = rest.split(")", 1)[0].strip()
            if not mid_raw:
                continue
            mid_l = mid_raw.lower()
            # Primeiro encontrado vence para manter determinismo.
            if mid_l not in mid:
                mid[mid_l] = f
            if "_" in mid_l:
                base_mid = mid_l.split("_", 1)[0].strip()
                if base_mid and base_mid not in mid:
                    mid[base_mid] = f
    except OSError:
        pass

    _FOTO_INDEX_CACHE[key] = (exato, mid)
    return exato, mid


def _indexar_fotos_recursivo(pasta: Path, prefixo: str) -> tuple[dict[str, Path], dict[str, Path]]:
    """Indexa JPG recursivamente (fallback robusto para ZIPs com pastas profundas)."""
    key = (str(pasta), (prefixo or "").strip().lower())
    cached = _FOTO_INDEX_RECURSIVO_CACHE.get(key)
    if cached is not None:
        return cached

    exato: dict[str, Path] = {}
    mid: dict[str, Path] = {}
    pref_l = (prefixo or "").strip().lower()
    start_l = f"{pref_l} ("

    try:
        for f in pasta.rglob("*"):
            if not f.is_file():
                continue
            name = f.name
            low = name.lower()
            if not low.endswith(".jpg"):
                continue
            if low not in exato:
                exato[low] = f
            if not low.startswith(start_l):
                continue
            rest = name[len(prefixo) + 2 :]
            if ")" not in rest:
                continue
            mid_raw = rest.split(")", 1)[0].strip()
            if not mid_raw:
                continue
            mid_l = mid_raw.lower()
            if mid_l not in mid:
                mid[mid_l] = f
            if "_" in mid_l:
                base_mid = mid_l.split("_", 1)[0].strip()
                if base_mid and base_mid not in mid:
                    mid[base_mid] = f
    except OSError:
        pass

    _FOTO_INDEX_RECURSIVO_CACHE[key] = (exato, mid)
    return exato, mid


def encontrar_foto_por_codigo_ou_numero(
    pasta: "Path",
    prefixo: str,
    codigo: str | int | None = None,
    numero: int | None = None,
) -> "Path | None":
    """
    Encontra arquivo de foto por código ou número.
    prefixo: "nc" ou "PDF". Procura em pasta/nc/ e pasta/PDF/ (ZIP extração) e na raiz.
    """
    pasta = Path(pasta)
    if not pasta.is_dir():
        return None
    prefix = f"{prefixo} ("
    suffix = ").jpg"
    bases = _pastas_busca_foto_extracao(pasta, prefixo)

    def _buscar_por_codigo_str(cod: str) -> "Path | None":
        cod = (cod or "").strip()
        if not cod:
            return None
        cod_l = cod.lower()
        for base in bases:
            exato, mid = _indexar_fotos_base(base, prefixo)
            hit = exato.get(f"{prefix}{cod}{suffix}".lower())
            if hit is not None:
                return hit
            hit = mid.get(cod_l)
            if hit is not None:
                return hit
        exato_r, mid_r = _indexar_fotos_recursivo(pasta, prefixo)
        hit = exato_r.get(f"{prefix}{cod}{suffix}".lower())
        if hit is not None:
            return hit
        hit = mid_r.get(cod_l)
        if hit is not None:
            return hit
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
        for cod in (str(n), str(n).zfill(5)):
            cod_l = cod.lower()
            for base in bases:
                exato, mid = _indexar_fotos_base(base, prefixo)
                hit = exato.get(f"{prefix}{cod}{suffix}".lower())
                if hit is not None:
                    return hit
                hit = mid.get(cod_l)
                if hit is not None:
                    return hit
            exato_r, mid_r = _indexar_fotos_recursivo(pasta, prefixo)
            hit = exato_r.get(f"{prefix}{cod}{suffix}".lower())
            if hit is not None:
                return hit
            hit = mid_r.get(cod_l)
            if hit is not None:
                return hit
    return None
