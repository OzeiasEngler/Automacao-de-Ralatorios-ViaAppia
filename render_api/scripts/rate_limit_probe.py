#!/usr/bin/env python3
"""
Script de referência para o rate limiting de rotas de probe (/.env, wp-login, etc.).
Mostra a configuração atual e simula quando o 429 seria retornado.

Uso:
  python -m render_api.scripts.rate_limit_probe              # imprime config
  python -m render_api.scripts.rate_limit_probe --simulate 35 # simula N requisições

Evolução futura (não urgente):
  - Se um IP bater em 3 probes diferentes (ex.: .env, .git, wp-admin) em menos de 10 s,
    marcar esse IP em cache (Redis ou dict em memória) como "malicioso".
  - Para IPs maliciosos: retornar 429 ou 403 em todas as rotas (inclusive legítimas)
    por 1 hora.
"""

import os
import sys


def _ler_int_env(name: str, default: int) -> int:
    try:
        v = os.environ.get(name, "").strip()
        return int(v) if v else default
    except ValueError:
        return default


def main():
    max_req = _ler_int_env("ARTESP_RATE_PROBE_MAX", 30)
    janela = _ler_int_env("ARTESP_RATE_PROBE_JANELA", 60)

    print("Rate limit (probe)")
    print("  ARTESP_RATE_PROBE_MAX   =", max_req, "(requisições máximas por IP)")
    print("  ARTESP_RATE_PROBE_JANELA=", janela, "(segundos)")
    print("  Rotas protegidas: /.env, /wp-login.php, /wp-admin, /.git/config, /config.php")
    print("  Ao exceder: 429 Too Many Requests + header Retry-After")
    print()

    if "--simulate" in sys.argv:
        idx = sys.argv.index("--simulate")
        n = int(sys.argv[idx + 1]) if idx + 1 < len(sys.argv) else 35
        if n <= max_req:
            print(f"Simulação: {n} requisições < {max_req} → todas retornariam 404 (probe bloqueado).")
        else:
            print(f"Simulação: {n} requisições > {max_req} → a partir da {max_req + 1}ª retornaria 429.")
    else:
        print("Use --simulate N para simular N requisições do mesmo IP.")


if __name__ == "__main__":
    main()
