# -*- coding: utf-8 -*-
"""
Conformidade: comparação Programação Mensal x Relatório de Executado.
- Déficit (Executado < Programado) → alerta "Item Pendente", vermelho no PDF.
- Superávit (Executado > Programado) → antecipação de cronograma.
- Gera PDF de alertas; prepara Excel limpo para ARTESP (preparar_excel_artesp).
- Saída dupla: ZIP protocolo (envio_artesp) e ZIP gestão (auditoria_interna) no app.
"""
from __future__ import annotations

import io
import logging
from pathlib import Path
from typing import Optional

import pandas as pd

logger = logging.getLogger(__name__)


def _normalizar_coluna(df: pd.DataFrame, nome: str, alternativas: list[str]) -> Optional[str]:
    """Retorna nome da coluna encontrada (case-insensitive, strip) ou None."""
    cols = {str(c).strip().lower(): c for c in df.columns}
    for alt in [nome.lower()] + [a.lower() for a in alternativas]:
        if alt in cols:
            return cols[alt]
    return None


def chave_linha(row: dict, col_item: str = "item", col_km: str = "km_inicial", col_sentido: str = "local") -> str:
    """Gera chave única para comparar mesmo lugar e serviço: item_km_sentido."""
    item = str(row.get(col_item, "") or "").strip()
    km = str(row.get(col_km, "") or "").strip()
    local = row.get(col_sentido)
    if isinstance(local, (list, tuple)) and local:
        sentido = str(local[0]).strip()
    else:
        sentido = str(local or "").strip()
    return f"{item}_{km}_{sentido}"


def analisar_conformidade(
    df_prog: pd.DataFrame,
    df_exec: pd.DataFrame,
    col_quantidade_plan: str = "quantidade",
    col_quantidade_real: str = "quantidade",
) -> pd.DataFrame:
    """
    Cruza programação e executado pela chave (item + km_inicial + sentido/local).
    Retorna comparativo com colunas: chave, qtd_plan, qtd_real, desvio, conformidade.
    conformidade: 'atraso' (déficit), 'antecipado' (superávit), 'conforme'.
    """
    def _chave_from_row(r, df_ref):
        item_c = _normalizar_coluna(df_ref, "item", ["item", "Item"]) or "item"
        km_c = _normalizar_coluna(df_ref, "km_inicial", ["km_inicial", "Km Inicial", "km"]) or "km_inicial"
        loc_c = _normalizar_coluna(df_ref, "local", ["local", "Local", "sentido", "Sentido"])
        d = r.to_dict()
        item = str(d.get(item_c, "") or "")
        km = str(d.get(km_c, "") or "")
        loc = d.get(loc_c)
        if isinstance(loc, (list, tuple)) and loc:
            sent = str(loc[0]).strip()
        else:
            sent = str(loc or "").strip()
        return f"{item}_{km}_{sent}"

    q_plan = _normalizar_coluna(df_prog, "quantidade", ["quantidade", "Quantidade", "qtd"])
    q_real = _normalizar_coluna(df_exec, "quantidade", ["quantidade", "Quantidade", "qtd", "executado"])
    if not q_plan:
        q_plan = df_prog.columns[0]
    if not q_real:
        q_real = df_exec.columns[0]

    df_prog = df_prog.copy()
    df_exec = df_exec.copy()
    df_prog["_chave"] = df_prog.apply(lambda r: _chave_from_row(r, df_prog), axis=1)
    df_exec["_chave"] = df_exec.apply(lambda r: _chave_from_row(r, df_exec), axis=1)

    agg_plan = df_prog.groupby("_chave", as_index=False).agg({q_plan: "sum"}).rename(columns={q_plan: "qtd_plan"})
    agg_real = df_exec.groupby("_chave", as_index=False).agg({q_real: "sum"}).rename(columns={q_real: "qtd_real"})

    comparativo = pd.merge(agg_plan, agg_real, on="_chave", how="left", suffixes=("", "_y"))
    comparativo["qtd_real"] = comparativo["qtd_real"].fillna(0)
    comparativo["desvio"] = comparativo["qtd_real"].astype(float) - comparativo["qtd_plan"].astype(float)

    def _conformidade(desvio):
        if desvio < 0:
            return "atraso"
        if desvio > 0:
            return "antecipado"
        return "conforme"

    comparativo["conformidade"] = comparativo["desvio"].apply(_conformidade)
    comparativo.rename(columns={"_chave": "chave"}, inplace=True)
    return comparativo


def gerar_pdf_alertas(
    comparativo: pd.DataFrame,
    lote: str,
    output_path: str | Path,
    titulo: str = "Relatório de Desempenho",
) -> Optional[str]:
    """
    Gera PDF com alertas (vermelho = déficit, verde = OK) e gráfico de barras
    comparando quantidade total por item. Usa reportlab.
    Retorna nome do arquivo gerado ou None.
    """
    try:
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.units import mm
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
        from reportlab.lib.enums import TA_CENTER, TA_LEFT
    except ImportError:
        logger.warning("reportlab não disponível — PDF de alertas não gerado.")
        return None

    path = Path(output_path)
    path.mkdir(parents=True, exist_ok=True)
    pdf_nome = f"Alerta_Programacao_Lote_{lote}.pdf"
    pdf_path = path / pdf_nome

    try:
        doc = SimpleDocTemplate(
            str(pdf_path),
            pagesize=A4,
            topMargin=15 * mm,
            bottomMargin=15 * mm,
            leftMargin=15 * mm,
            rightMargin=15 * mm,
        )
        styles = getSampleStyleSheet()
        style_titulo = ParagraphStyle(
            "TituloAlertas",
            parent=styles["Heading1"],
            fontSize=14,
            alignment=TA_CENTER,
            spaceAfter=4 * mm,
        )
        style_normal = ParagraphStyle("NormalAlertas", parent=styles["Normal"], fontSize=9, leading=11)
        elements = []

        elements.append(Paragraph(f"{titulo} — Lote {lote}", style_titulo))
        elements.append(Spacer(1, 3 * mm))

        # Tabela de itens com status
        dados = [["Chave", "Planejado", "Executado", "Desvio", "Status"]]
        for _, row in comparativo.iterrows():
            desvio = row.get("desvio", 0)
            if desvio < 0:
                status = f"ALERTA: Faltam {abs(desvio):.2f} un."
            elif desvio > 0:
                status = "Antecipado"
            else:
                status = "CONFORME"
            dados.append([
                str(row.get("chave", ""))[:40],
                str(row.get("qtd_plan", "")),
                str(row.get("qtd_real", "")),
                f"{desvio:.2f}",
                status,
            ])

        t = Table(dados, colWidths=[70 * mm, 25 * mm, 25 * mm, 25 * mm, 45 * mm])
        estilo_tabela = [
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0F2538")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f8f8f8")]),
        ]
        for i, row in comparativo.iterrows():
            if row.get("conformidade") == "atraso":
                idx = int(i) + 1
                if idx < len(dados):
                    estilo_tabela.append(("TEXTCOLOR", (0, idx), (-1, idx), colors.HexColor("#cc0000")))
                    estilo_tabela.append(("FONTNAME", (4, idx), (4, idx), "Helvetica-Bold"))
        t.setStyle(TableStyle(estilo_tabela))
        elements.append(t)
        elements.append(Spacer(1, 6 * mm))

        # Resumo por item (quantidade total) — texto simplificado (gráfico seria matplotlib + embed)
        elements.append(Paragraph("Resumo por item (total planejado x executado)", ParagraphStyle(
            "Sub", parent=styles["Heading2"], fontSize=11, spaceBefore=4 * mm, spaceAfter=2 * mm,
        )))
        por_item = comparativo.copy()
        por_item["item"] = por_item["chave"].str.split("_").str[0]
        resumo = por_item.groupby("item", as_index=False).agg({"qtd_plan": "sum", "qtd_real": "sum"})
        resumo["desvio"] = resumo["qtd_real"] - resumo["qtd_plan"]
        dados_resumo = [["Item", "Total Planejado", "Total Executado", "Desvio"]]
        for _, r in resumo.iterrows():
            dados_resumo.append([str(r["item"]), f"{r['qtd_plan']:.2f}", f"{r['qtd_real']:.2f}", f"{r['desvio']:.2f}"])
        t2 = Table(dados_resumo, colWidths=[40 * mm, 45 * mm, 45 * mm, 40 * mm])
        t2.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2c3e50")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 9),
            ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ]))
        elements.append(t2)

        # Gráfico de barras: quantidade total por item (programado x executado)
        try:
            from reportlab.graphics.shapes import Drawing
            from reportlab.graphics.charts.barcharts import VerticalBarChart
            from reportlab.lib.units import mm

            nomes = resumo["item"].astype(str).str[:20].tolist()
            if not nomes:
                nomes = ["—"]
                plan_list = [0]
                real_list = [0]
            else:
                plan_list = resumo["qtd_plan"].tolist()
                real_list = resumo["qtd_real"].tolist()
            max_val = max(max(plan_list or [0]), max(real_list or [0]), 1)

            largura_draw = 180 * mm
            altura_draw = 100 * mm
            d = Drawing(largura_draw, altura_draw)
            bc = VerticalBarChart()
            bc.x = 50
            bc.y = 20
            bc.width = largura_draw - 60
            bc.height = altura_draw - 40
            bc.data = [plan_list, real_list]
            bc.categoryAxis.categoryNames = nomes
            bc.categoryAxis.labels.angle = 45
            bc.categoryAxis.labels.fontSize = 7
            bc.valueAxis.valueMin = 0
            bc.valueAxis.valueMax = max_val * 1.1
            bc.bars[0].fillColor = colors.HexColor("#3498db")
            bc.bars[1].fillColor = colors.HexColor("#2ecc71")
            bc.barLabels.nudge = 2
            bc.barLabelFormat = "%.1f"
            d.add(bc, "")
            elements.append(Spacer(1, 3 * mm))
            elements.append(Paragraph(
                "Gráfico: Quantidade total por item (azul = planejado, verde = executado)",
                ParagraphStyle("LegendaChart", parent=styles["Normal"], fontSize=8, alignment=TA_CENTER),
            ))
            elements.append(Spacer(1, 2 * mm))
            elements.append(d)
        except Exception as chart_err:
            logger.debug("Gráfico de barras omitido: %s", chart_err)

        doc.build(elements)
        logger.info("PDF de alertas gerado: %s", pdf_path)
        return pdf_nome
    except Exception as e:
        logger.exception("Erro ao gerar PDF de alertas: %s", e)
        return None


def exportar_excel_auditoria(
    comparativo: pd.DataFrame,
    lote: str,
    output_path: str | Path,
) -> Optional[str]:
    """
    Exporta o comparativo (Programação x Executado) para Excel de auditoria interna.
    Arquivo independente da execução normal. Retorna nome do arquivo ou None.
    """
    if comparativo is None or comparativo.empty:
        return None
    path = Path(output_path)
    path.mkdir(parents=True, exist_ok=True)
    nome = f"Comparativo_Auditoria_Lote_{lote}.xlsx"
    arquivo = path / nome
    try:
        df_out = comparativo.copy()
        rename_map = {
            "chave": "Chave",
            "qtd_plan": "Planejado",
            "qtd_real": "Executado",
            "desvio": "Desvio",
            "conformidade": "Status",
        }
        df_out = df_out.rename(columns={k: v for k, v in rename_map.items() if k in df_out.columns})
        colunas = ["Chave", "Planejado", "Executado", "Desvio", "Status"]
        cols_finais = [c for c in colunas if c in df_out.columns]
        if not cols_finais:
            cols_finais = list(df_out.columns)[:10]
        df_out = df_out[cols_finais]
        df_out.to_excel(str(arquivo), index=False)
        return nome
    except Exception as e:
        logger.warning("Falha ao exportar Excel de auditoria: %s", e)
        return None


def mapa_conformidade_por_chave(comparativo: pd.DataFrame) -> dict:
    """Retorna dicionário chave -> conformidade ('atraso' | 'antecipado' | 'conforme')."""
    if comparativo is None or comparativo.empty or "chave" not in comparativo.columns:
        return {}
    return comparativo.set_index("chave")["conformidade"].to_dict()


def _chave_from_row_exec(r, df_ref: pd.DataFrame) -> str:
    """Gera chave item_km_sentido para uma linha do DataFrame de executado (igual ao usado em analisar_conformidade)."""
    item_c = _normalizar_coluna(df_ref, "item", ["item", "Item"]) or "item"
    km_c = _normalizar_coluna(df_ref, "km_inicial", ["km_inicial", "Km Inicial", "km"]) or "km_inicial"
    loc_c = _normalizar_coluna(df_ref, "local", ["local", "Local", "sentido", "Sentido"])
    d = r.to_dict()
    item = str(d.get(item_c, "") or "")
    km = str(d.get(km_c, "") or "")
    loc = d.get(loc_c)
    if isinstance(loc, (list, tuple)) and loc:
        sent = str(loc[0]).strip()
    else:
        sent = str(loc or "").strip()
    return f"{item}_{km}_{sent}"


# Colunas permitidas na versão ARTESP (protocolo externo) — sem status de erro/descumprimento
COLUNAS_ARTESP = ["Item", "KM_Inicial", "KM_Final", "Sentido", "Quantidade", "Data"]


def preparar_excel_artesp(
    df_exec: pd.DataFrame,
    comparativo: pd.DataFrame,
) -> pd.DataFrame:
    """
    Prepara DataFrame para envio ARTESP: apenas o que foi REALMENTE executado (Quantidade > 0),
    somente colunas técnicas (Item, KM_Inicial, KM_Final, Sentido, Quantidade, Data).
    Sem colunas de diferença, status ou alerta.
    """
    if df_exec is None or df_exec.empty or comparativo is None or comparativo.empty:
        return pd.DataFrame(columns=COLUNAS_ARTESP)
    if "chave" not in comparativo.columns or "qtd_real" not in comparativo.columns:
        return pd.DataFrame(columns=COLUNAS_ARTESP)

    chaves_executadas = set(
        comparativo.loc[comparativo["qtd_real"].astype(float) > 0, "chave"].astype(str).tolist()
    )
    if not chaves_executadas:
        return pd.DataFrame(columns=COLUNAS_ARTESP)

    df = df_exec.copy()
    df["_chave"] = df.apply(lambda r: _chave_from_row_exec(r, df_exec), axis=1)
    df = df[df["_chave"].isin(chaves_executadas)].copy()
    if df.empty:
        return pd.DataFrame(columns=COLUNAS_ARTESP)

    # Mapear colunas do df (normalizadas) para nomes ARTESP
    item_c = _normalizar_coluna(df, "item", ["item", "Item"])
    km_i = _normalizar_coluna(df, "km_inicial", ["km_inicial", "Km Inicial", "km"])
    km_f = _normalizar_coluna(df, "km_final", ["km_final", "Km Final"])
    loc_c = _normalizar_coluna(df, "local", ["local", "Local", "sentido", "Sentido"])
    qtd_c = _normalizar_coluna(df, "quantidade", ["quantidade", "Quantidade", "qtd", "executado"])
    data_c = _normalizar_coluna(df, "data_inicial", ["data_inicial", "Data Inicial", "data", "Data"])

    out = pd.DataFrame()
    out["Item"] = df[item_c] if item_c else ""
    out["KM_Inicial"] = df[km_i].astype(float) if km_i else 0
    out["KM_Final"] = df[km_f].astype(float) if km_f else df[km_i].astype(float) if km_i else 0
    if loc_c:
        out["Sentido"] = df[loc_c].apply(
            lambda x: x[0] if isinstance(x, (list, tuple)) and x else str(x) if pd.notna(x) else ""
        )
    else:
        out["Sentido"] = ""
    out["Quantidade"] = df[qtd_c].astype(float) if qtd_c else 0
    if data_c:
        out["Data"] = pd.to_datetime(df[data_c], errors="coerce").dt.strftime("%d/%m/%Y").fillna("")
    else:
        out["Data"] = ""
    out = out.dropna(how="all", subset=["Item", "KM_Inicial"])
    return out[COLUNAS_ARTESP] if not out.empty else pd.DataFrame(columns=COLUNAS_ARTESP)
