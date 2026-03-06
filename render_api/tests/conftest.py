"""Fixtures para testes da API."""
import os
import sys
from pathlib import Path

import pytest
from fastapi.testclient import TestClient


def pytest_configure(config):
    """Garante env de teste antes de qualquer import do app."""
    os.environ["ARTESP_WEB_USERS"] = '{"teste@artesp.local":"teste123"}'
    os.environ["ARTESP_ADMIN_EMAILS"] = "teste@artesp.local"
    os.environ["ARTESP_JWT_SECRET"] = "artesp-teste-jwt-secret-fixo-para-pytest"


# Garantir que o pacote está no path
_api_dir = Path(__file__).resolve().parent.parent
if str(_api_dir.parent) not in sys.path:
    sys.path.insert(0, str(_api_dir.parent))


def _get_app():
    try:
        import render_api.app as m
    except ImportError:
        import app as m
    return m.app, m


@pytest.fixture
def client():
    """Cliente de teste com override de auth (simula usuário logado)."""
    app, app_module = _get_app()

    def _mock_usuario():
        return "teste@artesp.local"

    app.dependency_overrides[app_module._get_usuario_autenticado] = _mock_usuario
    app.dependency_overrides[app_module._get_usuario_admin] = _mock_usuario

    try:
        yield TestClient(app)
    finally:
        app.dependency_overrides.clear()


@pytest.fixture
def client_sem_auth():
    """Cliente sem override - para testar 401 quando auth ausente."""
    app, _ = _get_app()
    return TestClient(app)


@pytest.fixture
def token(client):
    """Token JWT válido para testes."""
    r = client.post(
        "/auth/login",
        json={"email": "teste@artesp.local", "senha": "teste123"},
    )
    if r.status_code != 200:
        pytest.skip("Login falhou - verifique ARTESP_WEB_USERS")
    return r.json()["access_token"]


@pytest.fixture
def auth_headers(token):
    """Headers com Bearer token."""
    return {"Authorization": f"Bearer {token}"}
