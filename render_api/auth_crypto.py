#!/usr/bin/env python3
"""Funcoes de hash e verificacao de senha para autenticacao web ARTESP."""

import os
import re
import base64
import secrets
import hashlib

PBKDF2_PREFIX = "pbkdf2_sha256"


def _ler_int_env(nome, padrao):
    """
    Lê inteiro de variável de ambiente de forma tolerante.
    Aceita valores como:
      "600000"
      "600000 (recomendado)"
    """
    raw = (os.getenv(nome) or "").strip()
    if not raw:
        return padrao
    match = re.search(r"\d+", raw)
    if not match:
        return padrao
    try:
        return int(match.group(0))
    except Exception:
        return padrao


PBKDF2_DEFAULT_ITERATIONS = max(100000, _ler_int_env("ARTESP_WEB_PBKDF2_ITERATIONS", 390000))


def _b64decode_urlsafe_padded(value: str):
    padding = "=" * (-len(value) % 4)
    return base64.urlsafe_b64decode(value + padding)


def gerar_hash_senha(senha: str, iterations: int = PBKDF2_DEFAULT_ITERATIONS):
    """
    Gera hash seguro no formato:
      pbkdf2_sha256$<iteracoes>$<salt_b64>$<digest_b64>
    """
    if not senha:
        raise ValueError("Senha nao pode ser vazia para gerar hash.")
    if iterations < 100000:
        raise ValueError("Use pelo menos 100000 iteracoes para PBKDF2.")

    salt = os.urandom(16)
    digest = hashlib.pbkdf2_hmac("sha256", senha.encode("utf-8"), salt, iterations)
    salt_b64 = base64.urlsafe_b64encode(salt).decode("ascii").rstrip("=")
    digest_b64 = base64.urlsafe_b64encode(digest).decode("ascii").rstrip("=")
    return f"{PBKDF2_PREFIX}${iterations}${salt_b64}${digest_b64}"


def verificar_senha(senha_digitada: str, senha_armazenada: str):
    """
    Verifica senha nos formatos:
      1) Hash PBKDF2: pbkdf2_sha256$iter$salt_b64$digest_b64
      2) Texto puro (legado)
    """
    senha_digitada = senha_digitada or ""
    senha_armazenada = senha_armazenada or ""

    if senha_armazenada.startswith(f"{PBKDF2_PREFIX}$"):
        try:
            _, iter_str, salt_b64, digest_b64 = senha_armazenada.split("$", 3)
            iterations = int(iter_str)
            salt = _b64decode_urlsafe_padded(salt_b64)
            digest_ref = _b64decode_urlsafe_padded(digest_b64)
            digest_calc = hashlib.pbkdf2_hmac(
                "sha256", senha_digitada.encode("utf-8"), salt, iterations
            )
            return secrets.compare_digest(digest_calc, digest_ref)
        except Exception:
            return False

    return secrets.compare_digest(str(senha_armazenada), str(senha_digitada))
