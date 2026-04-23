# -*- coding: utf-8 -*-
"""Analisa GeoJSON gerado pelo exe: tamanho, pontos por feature, decimais."""
import os
import sys
import json

def main():
    if len(sys.argv) > 1:
        path = sys.argv[1]
    else:
        base = os.path.join(os.environ.get("USERPROFILE", ""), "OneDrive", "Ambiente de Trabalho", "ATeste Todos")
        candidatos = []
        for nome in ("RELATÓRIOS_ARTESP_GEOJSON_2026", "RELATORIOS_ARTESP_GEOJSON_2026"):
            for fn in ("L13_conservacao_2026_R0.geojson", "L13_conservacao_2026_r0.geojson"):
                candidatos.append(os.path.join(base, nome, "L13", "Anual", fn))
        path = None
        for p in candidatos:
            if os.path.exists(p):
                path = p
                break
        if path is None:
            path = candidatos[0] if candidatos else os.path.join(
                base, "RELATÓRIOS_ARTESP_GEOJSON_2026", "L13", "Anual", "L13_conservacao_2026_R0.geojson"
            )
    if not os.path.exists(path):
        print("Nao encontrado:", path)
        return

    size = os.path.getsize(path)
    size_mb = size / (1024 * 1024)
    print("Arquivo:", path)
    print("Tamanho: %.2f MB (%d bytes)" % (size_mb, size))

    with open(path, encoding="utf-8") as f:
        d = json.load(f)

    features = d.get("features", [])
    print("Features:", len(features))

    pts_por_feat = []
    for ft in features:
        g = ft.get("geometry")
        if not g:
            continue
        c = g.get("coordinates")
        if g.get("type") == "LineString" and c:
            pts_por_feat.append(len(c))
        elif g.get("type") == "MultiLineString" and c:
            pts_por_feat.append(sum(len(lin) for lin in c))

    if pts_por_feat:
        print("Pontos por geometria: min=%d max=%d media=%.0f" % (
            min(pts_por_feat), max(pts_por_feat), sum(pts_por_feat) / len(pts_por_feat)
        ))

    # Decimais na primeira coordenada
    for ft in features[:1]:
        g = ft.get("geometry")
        if not g or not g.get("coordinates"):
            continue
        coords = g["coordinates"]
        if g["type"] == "MultiLineString":
            coords = coords[0]
        if coords and len(coords[0]) >= 2:
            s = str(coords[0][0])
            dec = len(s.split(".")[-1]) if "." in s else 0
            print("Exemplo lon (decimais): %s -> %d decimais" % (s[:30], dec))
        break

    # Diagnóstico: redução aplicada?
    # step=2 + 4 decimais: média de pontos ~metade do que seria sem step; coords com 4 decimais
    if pts_por_feat and size_mb > 2.0:
        print("\n>>> GeoJSON grande (~%.1f MB). O exe pode estar com core antigo (sem redução)." % size_mb)
        print("    Recompile o .exe a partir de c:\\GeradorARTESP para usar step=2 e 4 decimais.")
    elif pts_por_feat:
        print("\n>>> Tamanho ok. Redução provavelmente aplicada (step=2, 4 decimais).")

if __name__ == "__main__":
    main()
