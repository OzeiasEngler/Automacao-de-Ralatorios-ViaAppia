"""Testa que salvar_geojson reduz tamanho quando ARTESP_GEOJSON_SIMPLIFY_STEP > 1."""
import os
import sys
import tempfile

# Garantir que o core do projeto está no path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..")))

import gerador_artesp_core as core


def _feature_com_muitos_pontos(n_pontos=200):
    """Uma LineString com n_pontos (simula trecho de rodovia)."""
    # São Paulo: lon -47..-46, lat -23..-22
    coords = [
        [-47.0 + i * 0.001, -23.0 + i * 0.0005]
        for i in range(n_pontos)
    ]
    return {
        "type": "Feature",
        "geometry": {"type": "LineString", "coordinates": coords},
        "properties": {
            "id": "L13_SP0000280_a.1.2_10.0_15.0_N_0001",
            "lote": "L13",
            "rodovia": "SP0000280",
            "item": "a.1.2",
            "km_inicial": 10.0,
            "km_final": 15.0,
        },
    }


def test_geojson_reducao_tamanho():
    """Com step=2 o GeoJSON salvo deve ser menor que com step=1."""
    obj = {
        "type": "FeatureCollection",
        "metadata": {"schema_version": "R0", "lote": "L13"},
        "crs": {"type": "name", "properties": {"name": "urn:ogc:def:crs:EPSG::4674"}},
        "features": [_feature_com_muitos_pontos(200) for _ in range(10)],
    }

    with tempfile.TemporaryDirectory() as tmp:
        path_sem_reducao = os.path.join(tmp, "sem_reducao.geojson")
        path_com_reducao = os.path.join(tmp, "com_reducao.geojson")

        # Salvar SEM redução (step=1)
        env_step = os.environ.get("ARTESP_GEOJSON_SIMPLIFY_STEP")
        os.environ["ARTESP_GEOJSON_SIMPLIFY_STEP"] = "1"
        try:
            core.salvar_geojson(path_sem_reducao, obj)
        finally:
            if env_step is not None:
                os.environ["ARTESP_GEOJSON_SIMPLIFY_STEP"] = env_step
            else:
                os.environ.pop("ARTESP_GEOJSON_SIMPLIFY_STEP", None)

        # Salvar COM redução (step=2)
        os.environ["ARTESP_GEOJSON_SIMPLIFY_STEP"] = "2"
        try:
            core.salvar_geojson(path_com_reducao, obj)
        finally:
            os.environ.pop("ARTESP_GEOJSON_SIMPLIFY_STEP", None)

        size_sem = os.path.getsize(path_sem_reducao)
        size_com = os.path.getsize(path_com_reducao)

    # step=2 mantém ~metade dos pontos → arquivo deve ser menor
    assert size_com < size_sem, (
        f"Esperado arquivo menor com step=2: sem_reducao={size_sem} bytes, com_reducao={size_com} bytes"
    )
    reducao_pct = (1 - size_com / size_sem) * 100
    print(f"OK: sem redução={size_sem:,} bytes, com step=2={size_com:,} bytes (redução ~{reducao_pct:.0f}%)")


if __name__ == "__main__":
    test_geojson_reducao_tamanho()
    print("Teste de redução GeoJSON: passou.")
