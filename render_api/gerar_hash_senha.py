#!/usr/bin/env python3
"""Utilitario para gerar hash PBKDF2 de senha para ARTESP_WEB_USERS."""

import argparse
import sys

try:
    # Quando executado como pacote: python -m render_api.gerar_hash_senha
    from .auth_crypto import PBKDF2_DEFAULT_ITERATIONS, gerar_hash_senha
except ImportError:
    # Quando executado direto no diretorio render_api
    from auth_crypto import PBKDF2_DEFAULT_ITERATIONS, gerar_hash_senha


def build_parser():
    parser = argparse.ArgumentParser(
        description="Gera hash PBKDF2 para uso no ARTESP_WEB_USERS."
    )
    parser.add_argument("senha", help="Senha em texto plano para converter em hash.")
    parser.add_argument(
        "--iterations",
        type=int,
        default=PBKDF2_DEFAULT_ITERATIONS,
        help=f"Numero de iteracoes PBKDF2 (padrao: {PBKDF2_DEFAULT_ITERATIONS}).",
    )
    parser.add_argument(
        "--email",
        help="Email opcional para imprimir no formato JSON usuario:hash.",
    )
    return parser


def main():
    args = build_parser().parse_args()

    try:
        senha_hash = gerar_hash_senha(args.senha, iterations=args.iterations)
    except Exception as e:
        print(f"[ERRO] {e}")
        return 1

    print("[OK] Hash gerado com sucesso.")
    print(senha_hash)
    if args.email:
        email = args.email.strip().lower()
        print("")
        print("Exemplo ARTESP_WEB_USERS (JSON):")
        print(f'{{"{email}":"{senha_hash}"}}')
    return 0


if __name__ == "__main__":
    sys.exit(main())
