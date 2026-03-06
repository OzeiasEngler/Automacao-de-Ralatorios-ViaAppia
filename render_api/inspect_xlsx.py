# -*- coding: utf-8 -*-
"""
Inspeciona planilha ARTESP (ex.: L21 Programação Anual) para geracao GeoJSON.
Uso: python inspect_xlsx.py "C:/caminho/para/L21 - Programação Anual 2026.xlsx"
     ou arraste o arquivo .xlsx para o script.
"""
import pandas as pd
import sys
import os

# Caminho padrao: tenta workspace primeiro
DEFAULT_PATH = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "data", "anual", "ANUAL_2026.xlsx"
)
if not os.path.isfile(DEFAULT_PATH):
    DEFAULT_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "L21 - Programação Anual 2026.xlsx")


def norm_key(s):
    s = (s or "").strip().lower()
    for c in " \n\r_-":
        s = s.replace(c, "")
    return s


def main():
    path = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_PATH
    if not os.path.isfile(path):
        print("Arquivo nao encontrado:", path)
        print("Uso: python inspect_xlsx.py \"caminho\\para\\planilha.xlsx\"")
        sys.exit(1)

    print("Planilha:", path)
    print()

    try:
        df_raw = pd.read_excel(path, sheet_name=0, header=None)
    except Exception as e:
        print("Erro ao ler Excel:", e)
        sys.exit(1)

    print("Shape:", df_raw.shape[0], "linhas x", df_raw.shape[1], "colunas")
    print()

    # Linha 0 = cabecalho (como no gerador_artesp_core.ler_excel)
    col_names = df_raw.iloc[0].values
    print("--- Colunas (linha 1 do Excel) ---")
    for i, v in enumerate(col_names):
        key = norm_key(str(v))
        print("  {:2d}: {!r}  -> norm_key: {!r}".format(i, v, key[:50] if key else ""))

    # Mapeamento esperado pelo core (normalizar_colunas_df)
    mapa = {}
    for c in col_names:
        ck = norm_key(str(c))
        if "lote" in ck:
            mapa[str(c)] = "lote"
        elif "rodovia" in ck or ck == "rod":
            mapa[str(c)] = "rodovia"
        elif "programa" in ck:
            mapa[str(c)] = "programa"
        elif "subitem" in ck:
            mapa[str(c)] = "subitem"
        elif "item" in ck and "sub" not in ck and "detalh" not in ck:
            mapa[str(c)] = "item"
        elif "detalhamento" in ck or "descricao" in ck or "servico" in ck:
            mapa[str(c)] = "detalhamento_servico"
        elif "unidade" in ck or "unid" in ck or ck == "un":
            mapa[str(c)] = "unidade"
        elif "quantidade" in ck or "qtd" in ck or "qtde" in ck:
            mapa[str(c)] = "quantidade"
        elif "kminicial" in ck or "kminicio" in ck or "kmini" in ck:
            mapa[str(c)] = "km_inicial"
        elif "kmfinal" in ck or "kmfim" in ck:
            mapa[str(c)] = "km_final"
        elif "local" in ck or "pista" in ck:
            mapa[str(c)] = "local"
        elif "datainicial" in ck or "datainicio" in ck or "dataini" in ck:
            mapa[str(c)] = "data_inicial"
        elif "datafinal" in ck or "datafim" in ck:
            mapa[str(c)] = "data_final"
        elif "observ" in ck or "obs" in ck:
            mapa[str(c)] = "observacoes_gerais"
        elif "latitude" in ck or ck == "lat":
            mapa[str(c)] = "Latitude"
        elif "longitude" in ck or ck in ("lon", "lng"):
            mapa[str(c)] = "Longitude"

    print()
    print("--- Colunas reconhecidas pelo gerador (nome interno) ---")
    for orig, nome in sorted(mapa.items(), key=lambda x: x[1]):
        print("  ", nome, "<-", orig[:45])

    obrigatorias = ["lote", "rodovia", "item", "detalhamento_servico", "unidade", "quantidade",
                    "km_inicial", "km_final", "local", "data_inicial", "data_final", "observacoes_gerais"]
    faltando = [o for o in obrigatorias if o not in mapa.values()]
    if faltando:
        print()
        print("ATENCAO: colunas obrigatorias nao reconhecidas:", faltando)
    else:
        print()
        print("OK: todas as colunas obrigatorias reconhecidas.")

    # Dados a partir da linha 6 (linha_inicio_dados = 6 no core)
    df = df_raw.iloc[5:].copy()
    df.columns = col_names
    df = df.rename(columns=mapa)
    df = df.dropna(how="all")
    print()
    print("--- Dados: primeira linha de dados (apos linha 6) ---")
    if len(df) > 0:
        print(df.head(1).to_string(max_colwidth=40))
        if "lote" in df.columns:
            unicos = df["lote"].dropna().astype(str).str.strip().unique()
            print()
            print("Valores unicos na coluna 'lote':", list(unicos)[:20])
    print()
    print("Total de linhas de dados (apos remover vazias):", len(df))

    # --- Rodovias existentes na planilha ---
    if len(df) > 0 and "rodovia" in df.columns:
        import re
        def normalizar_rodovia(nome):
            if pd.isna(nome):
                return ""
            limpo = str(nome).upper().replace("-", "").replace(" ", "").replace("/", "").strip()
            limpo = limpo.replace("HRT", "HTR")
            if "SPD" in limpo:
                m = re.match(r"(\d{1,2})SPD(\d+)", limpo)
                if m:
                    return f"{int(m.group(1)):02d}SPD{int(m.group(2)):06d}"
            if limpo.startswith("SPM"):
                m = re.match(r"SPM(\d+)([A-Z])?", limpo)
                if m:
                    return f"SPM{int(m.group(1)):05d}{m.group(2) or ''}"
            if limpo.startswith("SPI"):
                m = re.match(r"SPI(\d+)([A-Z])?", limpo)
                if m:
                    return f"SPI{int(m.group(1)):06d}{m.group(2) or ''}"
            if limpo.startswith("SP") and not limpo.startswith(("SPA", "SPI", "SPM", "SPD")):
                nums = "".join(filter(str.isdigit, limpo))
                return f"SP{int(nums):07d}" if nums else limpo
            return limpo

        rodovias_planilha = df["rodovia"].dropna().astype(str).apply(normalizar_rodovia)
        rodovias_planilha = rodovias_planilha[rodovias_planilha.str.len() > 0].unique().tolist()
        rodovias_planilha = sorted(set(rodovias_planilha))

        print()
        print("=== RODOVIAS NA PLANILHA (normalizadas) ===")
        print("Total:", len(rodovias_planilha))
        for r in rodovias_planilha:
            print(" ", r)

        # Comparar com L21 (malha)
        try:
            sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
            import gerador_artesp_core as core
            rodovias_l21 = set(core.RODOVIAS_POR_LOTE.get("L21", []))
            na_malha = [r for r in rodovias_planilha if r in rodovias_l21]
            fora_malha = [r for r in rodovias_planilha if r not in rodovias_l21]
            print()
            print("--- Comparacao com a malha do Lote 21 ---")
            print("Rodovias da planilha que ESTAO na malha L21:", len(na_malha))
            for r in sorted(na_malha):
                print("  [OK]", r)
            if fora_malha:
                print()
                print("Rodovias da planilha que NAO estao na malha L21:", len(fora_malha))
                for r in sorted(fora_malha):
                    print("  [FORA]", r)
        except Exception as e:
            print("(Nao foi possivel comparar com core:", e, ")")


if __name__ == "__main__":
    main()
