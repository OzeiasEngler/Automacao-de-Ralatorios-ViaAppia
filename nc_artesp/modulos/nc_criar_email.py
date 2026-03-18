"""
modulos/nc_criar_email.py
────────────────────────────────────────────────────────────────────────────
Equivalente VBA: NC_Artesp_Criar_Email_Pad_25

A partir de e-mails selecionados no Outlook + arquivos XLS em Exportar/:

  1. Lê cada XLS individual da pasta Exportar/ (saída do mod 01).
  2. Coluna U = responsável pelo apontamento (fiscal que fez o apontamento); esse é quem recebe a resposta.
  3. Para cada planilha: identifica o e-mail selecionado cujo remetente corresponde
     ao responsável (U) e cria um único ReplyAll para esse par planilha↔fiscal.
     Assim, múltiplas planilhas (fiscais diferentes) geram respostas cada uma ao correto.
  4. Assunto/Corpo/CC como antes; salva rascunho (não envia automaticamente).

Colunas do XLS (planilha EAF individual, saída M01 – conforme macros):
  Utilizadas: C,D,E,F,H,I,J,K,L,Q,T,U,V,P. U=responsável (fiscal). P=grupo EAF. C=código fiscalização.
  Na criação do e-mail: verifica o código (e grupo/trecho) e atribui o destinatário responsável automaticamente
  via MAPA_EAF[].email (grupo da col P ou grupo obtido por rodovia+km); fallback col U se contiver @.
  Não utilizadas na extração: A,B,G,M,N,O,R,S,W,X,Y... (ver docs/ANALISE_COLUNAS_MODELO_KCOR_KRIA.md)

Imagem embutida: pdf(N).jpg com CID = "pdf%20(N).jpg"
Dimensão HTML: height=295 width=711

Requer: Windows + Outlook instalado + pywin32
"""

import logging
from pathlib import Path

from openpyxl import load_workbook

from config import M01_EXPORTAR, M02_FOTOS_PDF, MAPA_EAF, NC_EMAIL_CC, RODOVIAS
from utils.helpers import encontrar_foto_por_codigo_ou_numero, obter_grupo_empresa_por_trecho, path_foto_pdf

logger = logging.getLogger(__name__)

# Índices de coluna (1-based) – planilha EAF individual
_C  = 3   # código fiscalização
_D  = 4   # data fiscalização
_E  = 5   # horário
_F  = 6   # rodovia
_G  = 7   # concessionária
_H  = 8   # km inicial
_I  = 9   # m inicial
_J  = 10  # km final
_K  = 11  # m final
_L  = 12  # sentido
_M  = 13  # data retorno
_N  = 14  # status retorno
_O  = 15  # tipo atividade
_P  = 16  # grupo atividade
_Q  = 17  # atividade
_R  = 18  # nº notificação
_S  = 19  # data envio
_T  = 20  # data reparo
_U  = 21  # responsável
_V  = 22  # nº foto

_LINHA_INICIO = 5


def _cell(ws, row: int, col: int) -> str:
    v = ws.cell(row=row, column=col).value
    return str(v).strip() if v is not None else ""


def _normalizar_rodovia(valor: str) -> str:
    prefixos = {"SP 075": "SP 075", "SP 127": "SP 127",
                "SP 280": "SP 280", "SP 300": "SP 300",
                "SPI 10": "SPI 102/300"}
    for k, v in prefixos.items():
        if valor.startswith(k):
            return v
    return valor[:6].strip()


def _ler_xls(arq: Path) -> list[dict]:
    """Lê todas as NCs de um XLS individual (saída mod 01)."""
    wb = load_workbook(str(arq), data_only=True)
    # Tenta planilha "Sheet0" primeiro (formato VBA), depois ativa
    ws = None
    for nome in wb.sheetnames:
        if nome.lower() in ("sheet0", "sheet1", "planilha1", "folha1"):
            ws = wb[nome]
            break
    if ws is None:
        ws = wb.active

    ultima = ws.max_row
    for r in range(ultima, _LINHA_INICIO - 1, -1):
        if ws.cell(row=r, column=_C).value:
            ultima = r
            break

    ncs = []
    for r in range(_LINHA_INICIO, ultima + 1):
        rod_raw = _cell(ws, r, _F)
        ncs.append({
            "cod":         _cell(ws, r, _C),
            "data_fisc":   _cell(ws, r, _D),
            "rodovia":     _normalizar_rodovia(rod_raw),
            "km_i":        _cell(ws, r, _H),
            "m_i":         _cell(ws, r, _I),
            "sentido":     _cell(ws, r, _L),
            "atividade":   _cell(ws, r, _Q),
            "data_rep":    _cell(ws, r, _T),
            "responsavel": _cell(ws, r, _U),  # responsável pelo apontamento (fiscal) — col U
            "foto":        _cell(ws, r, _V),
            "grupo":       _cell(ws, r, _P),  # grupo EAF (col P) — para obter e-mail do MAPA_EAF
        })
    wb.close()
    return [n for n in ncs if n["atividade"]]


def _resolver_foto_pdf(pasta_fotos_pdf: Path, nc: dict) -> "tuple[Path | None, str]":
    """
    Resolve o arquivo da foto PDF para uma NC, alinhado à lógica de renomeação do extrator.
    O extrator grava por código de fiscalização (ex.: PDF (896643).jpg, PDF (NC.13.1039).jpg);
    a planilha tem col C = código e col V = número da foto.
    Retorna (path do arquivo ou None, valor CID para Content-ID — nome do arquivo com espaço → %20).
    """
    if not pasta_fotos_pdf or not Path(pasta_fotos_pdf).is_dir():
        return (None, "")
    pasta = Path(pasta_fotos_pdf)
    cod = (nc.get("cod") or "").strip()
    foto_raw = nc.get("foto") or ""
    numero = None
    try:
        numero = int(float(str(foto_raw).strip()))
    except (ValueError, TypeError):
        pass
    # Tentar path exato por número (compatível com PDF (1).jpg, PDF (00001).jpg)
    if numero is not None:
        p = path_foto_pdf(pasta, numero)
        if p.is_file():
            cid = p.name.replace(" ", "%20")
            return (p, cid)
        p5 = path_foto_pdf(pasta, str(numero).zfill(5))
        if p5.is_file():
            cid = p5.name.replace(" ", "%20")
            return (p5, cid)
    # Buscar por código ou número (ex.: PDF (896643).jpg, PDF (NC.13.1039).jpg, variantes _1, etc.)
    encontrado = encontrar_foto_por_codigo_ou_numero(
        pasta, "PDF", codigo=cod if cod else None, numero=numero
    )
    if encontrado and encontrado.is_file():
        cid = encontrado.name.replace(" ", "%20")
        return (encontrado, cid)
    return (None, "")


def _bloco_html_nc(nc: dict, cid: str | None = None) -> str:
    """Gera o bloco HTML de uma NC (cabeçalho + imagem CID). cid: se informado, usa no img; senão pdf%20({foto}).jpg."""
    if not cid:
        cid = f"pdf%20({nc['foto']}).jpg"
    cab = (
        f"{nc['rodovia']} - km {nc['km_i']},{nc['m_i']} {nc['sentido']} "
        f"- Const: {nc['data_fisc']} - Prazo: {nc['data_rep']} "
        f"- {nc['atividade']} - Cod. Fisc.: {nc['cod']}"
    )
    return (
        f"<b><u>{cab}</u></b><BR><BR>"
        f'<img src="cid:{cid}" height=295 width=711>'
        "<BR><BR><BR><BR>"
    )


def _html_saudacao() -> str:
    return (
        "Prezados,<BR><BR>"
        "Seguem registros fotográficos das superações de não conformidade, "
        "dentro do prazo regulamentado.<BR><BR>"
    )


def _assunto_enriquecido(assunto_original: str, nc_ref: dict) -> str:
    """Monta o assunto enriquecido (ReplyAll + dados da NC)."""
    base = assunto_original.replace(" [Email Externo] ", "")
    return (
        f"{base} - {nc_ref['rodovia']} ({nc_ref['atividade']}) "
        f"- Const: {nc_ref['data_fisc']} - Prazo: {nc_ref['data_rep']}"
    )


def _cc_str() -> str:
    return ";".join(NC_EMAIL_CC)


def _email_por_grupo(grupo_val, mapa_eaf: list) -> str:
    """Retorna o e-mail do responsável do grupo EAF (MAPA_EAF[].email), se definido."""
    if not mapa_eaf:
        return ""
    if grupo_val is None or (isinstance(grupo_val, (int, float)) and grupo_val == 0):
        return ""
    try:
        g = int(float(str(grupo_val).strip()))
    except (ValueError, TypeError):
        return ""
    if g == 0:
        return ""
    for entry in mapa_eaf:
        if entry.get("grupo") == g:
            email = (entry.get("email") or "").strip()
            if email and "@" in email:
                return email
            break
    return ""


def _km_float(km_cel: str) -> float | None:
    """Converte célula de km (ex. 143+800 ou 143.8) para float."""
    if not km_cel:
        return None
    s = str(km_cel).strip().replace(",", ".")
    if "+" in s:
        parts = s.split("+", 1)
        try:
            return float(parts[0].strip()) + float(parts[1].strip()) / 1000.0
        except (ValueError, TypeError):
            return None
    try:
        return float(s)
    except (ValueError, TypeError):
        return None


def _destinatario_responsavel_automatico(nc_ref: dict, mapa_eaf: list | None = None) -> str:
    """
    Atribui o destinatário responsável automaticamente a partir dos dados da NC (código/grupo).
    Na criação do e-mail, verifica o código (e grupo/rodovia+km) e resolve o e-mail do responsável:
    1) E-mail do grupo EAF: col P (grupo) ou grupo obtido por rodovia+km → MAPA_EAF[].email.
    2) Senão, col U (responsável) se contiver @.
    Assim o To é definido pelo grupo da fiscalização (código/trecho), sem depender só da coluna U.
    """
    mapa = mapa_eaf if mapa_eaf is not None else MAPA_EAF
    cod = nc_ref.get("cod") or ""
    email = _email_por_grupo(nc_ref.get("grupo"), mapa)
    if email:
        logger.debug("Destinatário automático (código %s, grupo col P): %s", cod, email)
        return email
    grupo, empresa = obter_grupo_empresa_por_trecho(
        nc_ref.get("rodovia"), _km_float(nc_ref.get("km_i")), mapa
    )
    email = _email_por_grupo(grupo, mapa)
    if email:
        logger.debug("Destinatário automático (código %s, grupo %s %s por trecho): %s", cod, grupo, empresa, email)
        return email
    resp = (nc_ref.get("responsavel") or "").strip()
    if resp and "@" in resp:
        return resp
    return ""


def _responsavel_casa_com_remetente(responsavel: str, sender_email: str, sender_name: str) -> bool:
    """
    Verifica se o responsável pelo apontamento (col U da planilha) corresponde ao remetente do e-mail.
    Aceita: e-mail igual, nome contido no remetente, ou e-mail contido no responsável.
    """
    r = (responsavel or "").strip().lower()
    if not r:
        return False
    se = (sender_email or "").strip().lower()
    sn = (sender_name or "").strip()
    if se and r == se:
        return True
    if sn and r in sn.lower():
        return True
    if se and se in r:
        return True
    if r and r in se:
        return True
    return False


def _criar_via_outlook(pasta_xls: Path,
                        pasta_fotos_pdf: Path,
                        callback_progresso=None) -> int:
    """
    Para cada XLS, identifica o fiscal responsável (col U) e cria o reply no e-mail
    selecionado cujo remetente corresponda a esse fiscal. Assim cada planilha gera
    resposta para o fiscal correto (múltiplas planilhas = múltiplos fiscais).
    Retorna número de e-mails rascunhados.
    """
    try:
        import win32com.client as win32
    except ImportError:
        raise ImportError(
            "pywin32 não instalado. Execute: pip install pywin32\n"
            "Requer Windows + Outlook instalado e aberto."
        )

    # Coletar XLS
    arquivos = sorted([
        f for f in pasta_xls.glob("*.xls*")
        if not f.name.startswith("~") and not f.name.startswith("_")
    ])
    if not arquivos:
        logger.warning(f"Nenhum XLS em: {pasta_xls}")
        return 0

    outlook   = win32.Dispatch("Outlook.Application")
    explorer  = outlook.ActiveExplorer()
    selection = explorer.Selection

    # Lista de e-mails selecionados com remetente (para casar com col U)
    itens_selecao = []
    for i in range(1, selection.Count + 1):
        item = selection.Item(i)
        if item.Class != 43:  # 43 = olMail
            continue
        try:
            sender_email = getattr(item, "SenderEmailAddress", "") or ""
            sender_name  = getattr(item, "SenderName", "") or ""
            itens_selecao.append((item, sender_email, sender_name))
        except Exception as e:
            logger.warning(f"  E-mail ignorado (erro ao obter remetente): {e}")
    if not itens_selecao:
        logger.warning("Nenhum e-mail de correio selecionado no Outlook.")
        return 0

    PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
    rascunhos = 0

    for idx, arq in enumerate(arquivos):
        if callback_progresso:
            callback_progresso(idx + 1, len(arquivos), f"Processando: {arq.name[:50]}")

        ncs = _ler_xls(arq)
        if not ncs:
            logger.warning(f"  Nenhuma NC em {arq.name}")
            continue

        nc_ref = ncs[0]
        responsavel = (nc_ref.get("responsavel") or "").strip()

        # Encontrar o e-mail selecionado cujo remetente corresponde ao responsável pelo apontamento (col U)
        item_correspondente = None
        for item, sender_email, sender_name in itens_selecao:
            if _responsavel_casa_com_remetente(responsavel, sender_email, sender_name):
                item_correspondente = item
                break

        if item_correspondente is None:
            logger.warning(
                f"  Planilha {arq.name}: responsável pelo apontamento (col U) '{responsavel or '(vazio)'}' não corresponde a nenhum remetente dos e-mails selecionados. "
                "Rascunho não criado. Verifique a coluna U (responsável) e selecione os e-mails dos fiscais corretos."
            )
            continue

        reply = item_correspondente.ReplyAll()
        reply.Subject = _assunto_enriquecido(reply.Subject, nc_ref)
        reply.CC = _cc_str()
        to_addr = _destinatario_responsavel_automatico(nc_ref)
        reply.To = to_addr if to_addr else ""

        corpo_html = _html_saudacao()
        for nc in ncs:
            foto_path, cid = _resolver_foto_pdf(pasta_fotos_pdf, nc)
            if not foto_path or not foto_path.exists():
                logger.warning(f"  Foto não encontrada para NC cod={nc.get('cod')} foto={nc.get('foto')}: PDF (cod).jpg ou PDF (n).jpg")
                continue
            attach = reply.Attachments.Add(str(foto_path))
            pa = attach.PropertyAccessor
            pa.SetProperty(PR_ATTACH_CONTENT_ID, cid)
            corpo_html += _bloco_html_nc(nc, cid=cid)

        olFormatHTML = 2
        if reply.BodyFormat == olFormatHTML:
            reply.HTMLBody = f"<p>{corpo_html}</p>" + reply.HTMLBody
        else:
            reply.Body = corpo_html.replace("<BR>", "\n").replace("<b><u>", "").replace("</u></b>", "") + reply.Body

        reply.Save()
        rascunhos += 1
        logger.info(f"  Rascunho salvo (fiscal: {responsavel or 'N/A'}): {reply.Subject[:60]}")

    return rascunhos


def _criar_eml(pasta_xls: Path,
               pasta_fotos_pdf: Path,
               pasta_saida: Path,
               callback_progresso=None) -> list[Path]:
    """
    Gera arquivos .eml para cada XLS (pode ser aberto em qualquer cliente de e-mail).
    Usa base64 para as imagens (sem CID inline, mas portátil).
    """
    import base64
    import email.mime.multipart as mm
    import email.mime.text      as mt
    import email.mime.image     as mi
    from email.utils import make_msgid
    from utils.helpers import garantir_pasta, timestamp_agora

    garantir_pasta(pasta_saida)

    arquivos = sorted([
        f for f in pasta_xls.glob("*.xls*")
        if not f.name.startswith("~") and not f.name.startswith("_")
    ])
    gerados = []

    for idx, arq in enumerate(arquivos):
        if callback_progresso:
            callback_progresso(idx + 1, len(arquivos), f"Gerando .eml: {arq.name[:50]}")

        ncs = _ler_xls(arq)
        if not ncs:
            continue

        nc_ref = ncs[0]
        responsavel = (nc_ref.get("responsavel") or "").strip()
        assunto = _assunto_enriquecido("RE: Apontamento NC Artesp", nc_ref)

        msg = mm.MIMEMultipart("related")
        msg["Subject"] = assunto
        msg["From"]    = "artesp.nc@conservacao.br"
        msg["CC"]      = _cc_str()
        to_addr = _destinatario_responsavel_automatico(nc_ref)
        if to_addr:
            msg["To"] = to_addr

        corpo_html = _html_saudacao()
        partes_img = []

        for nc in ncs:
            foto_path, _ = _resolver_foto_pdf(pasta_fotos_pdf, nc)
            if not foto_path or not foto_path.exists():
                continue
            cid = make_msgid(domain="artesp.nc")
            cid_limpo = cid[1:-1]  # sem < >
            corpo_html += (
                f"<b><u>{nc['rodovia']} - km {nc['km_i']},{nc['m_i']} {nc['sentido']}"
                f" - {nc['atividade']}</u></b><BR><BR>"
                f'<img src="cid:{cid_limpo}" height=295 width=711><BR><BR><BR><BR>'
            )
            partes_img.append((foto_path, cid))

        html_part = mt.MIMEText(f"<html><body>{corpo_html}</body></html>", "html", "utf-8")
        msg.attach(html_part)

        for foto_path, cid in partes_img:
            with open(str(foto_path), "rb") as f:
                img = mi.MIMEImage(f.read(), _subtype="jpeg")
                img.add_header("Content-ID", cid)
                img.add_header("Content-Disposition", "inline",
                               filename=foto_path.name)
                msg.attach(img)

        nome_eml = f"{timestamp_agora()} - {arq.stem}.eml"
        destino  = pasta_saida / nome_eml
        destino.write_text(msg.as_string(), encoding="utf-8")
        gerados.append(destino)
        logger.info(f"  .eml gerado: {destino.name}")

    return gerados


def executar(pasta_xls: Path | None = None,
             pasta_fotos_pdf: Path | None = None,
             usar_outlook: bool = True,
             pasta_saida_eml: Path | None = None,
             callback_progresso=None) -> dict:
    """
    Cria e-mails padrão de resposta NC.

    usar_outlook=True  → ReplyAll via Outlook COM (requer Outlook aberto
                          com e-mail(s) selecionado(s))
    usar_outlook=False → gera arquivos .eml portáteis

    Retorna dict com 'rascunhos' (int) e/ou 'eml' (list[Path]).
    """
    pasta_xls       = pasta_xls       or M01_EXPORTAR
    pasta_fotos_pdf = pasta_fotos_pdf or M02_FOTOS_PDF

    resultado = {"rascunhos": 0, "eml": []}

    if usar_outlook:
        logger.info("Módulo NC Email: criando rascunhos via Outlook...")
        rascunhos = _criar_via_outlook(pasta_xls, pasta_fotos_pdf, callback_progresso)
        resultado["rascunhos"] = rascunhos
        logger.info(f"Módulo NC Email concluído: {rascunhos} rascunho(s).")
    else:
        logger.info("Módulo NC Email: gerando arquivos .eml...")
        from config import M04_SAIDA
        pasta_eml = pasta_saida_eml or (M04_SAIDA / "emails")
        emls = _criar_eml(pasta_xls, pasta_fotos_pdf, pasta_eml, callback_progresso)
        resultado["eml"] = emls
        logger.info(f"Módulo NC Email concluído: {len(emls)} arquivo(s) .eml.")

    if callback_progresso:
        callback_progresso(1, 1, "Módulo NC Email concluído.")

    return resultado
