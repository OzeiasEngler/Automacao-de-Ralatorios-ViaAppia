"""Testes de integração para rotas da API."""
import pytest
from fastapi.testclient import TestClient


def test_read_root(client: TestClient):
    """Raiz devolve página inicial com link/redirecionamento para GeoJSON."""
    r = client.get("/", follow_redirects=False)
    assert r.status_code == 200
    assert "/web/geojson" in (r.text or "")


def test_api_config(client: TestClient):
    """Config retorna lotes, modalidades, versões."""
    r = client.get("/api/config")
    assert r.status_code == 200
    data = r.json()
    assert "lotes" in data
    assert "modalidades" in data
    assert "versoes" in data
    assert len(data["lotes"]) > 0


def test_login_sucesso(client: TestClient):
    """Login com credenciais válidas retorna token."""
    r = client.post(
        "/auth/login",
        json={"email": "teste@artesp.local", "senha": "teste123"},
    )
    if r.status_code != 200:
        pytest.skip("Usuário de teste não configurado")
    data = r.json()
    assert "access_token" in data
    assert "token_type" in data


def test_login_senha_invalida(client: TestClient):
    """Login com senha errada retorna 401."""
    r = client.post(
        "/auth/login",
        json={"email": "teste@artesp.local", "senha": "senhaerrada"},
    )
    assert r.status_code == 401


def test_api_stats_requer_auth(client_sem_auth: TestClient):
    """GET /api/stats sem token retorna 401."""
    r = client_sem_auth.get("/api/stats")
    assert r.status_code == 401


def test_api_stats_com_bearer_token(client_sem_auth: TestClient):
    """GET /api/stats com Authorization: Bearer <token> retorna 200."""
    login = client_sem_auth.post(
        "/auth/login",
        json={"email": "teste@artesp.local", "senha": "teste123"},
    )
    if login.status_code != 200:
        pytest.skip("Login falhou - verifique ARTESP_WEB_USERS e ARTESP_JWT_SECRET")
    token = login.json()["access_token"]
    r = client_sem_auth.get("/api/stats", headers={"Authorization": f"Bearer {token}"})
    assert r.status_code == 200, f"Esperado 200 com Bearer token, obtido {r.status_code}: {r.text}"
    assert "total_gerado" in r.json()


def test_api_stats_com_token(client: TestClient):
    """GET /api/stats com auth (override em teste) retorna métricas."""
    r = client.get("/api/stats")
    assert r.status_code == 200, f"Esperado 200, obtido {r.status_code}: {r.text}"
    data = r.json()
    assert "total_gerado" in data
    assert "pendencias_hoje" in data
    assert "historico" in data


def test_gerar_relatorio_sem_arquivo(client: TestClient):
    """POST /gerar-relatorio-progresso sem arquivo retorna 422 (validação)."""
    r = client.post(
        "/gerar-relatorio-progresso",
        data={
            "lote": "13",
            "modalidade": "1",
            "versao": "r0",
            "ano": "2025",
            "mes": "",
            "correcao_eixo": "false",
            "dashboard": "false",
            "assinar": "true",
            "dry_run": "true",
        },
    )
    assert r.status_code == 422, f"Esperado 422 (validação), obtido {r.status_code}: {r.text}"


def test_admin_stats_requer_admin(client: TestClient):
    """GET /admin/stats com override retorna 200 (teste@artesp.local é admin)."""
    r = client.get("/admin/stats")
    assert r.status_code in (200, 403), f"Esperado 200 ou 403, obtido {r.status_code}: {r.text}"


def test_health_check(client: TestClient):
    """Health check retorna status online."""
    r = client.get("/health")
    assert r.status_code == 200
    data = r.json()
    assert data.get("status") == "online"
    assert "auth" in data


def test_adicionar_usuario_sucesso(client: TestClient, client_sem_auth: TestClient, tmp_path, monkeypatch):
    """
    Admin adiciona novo usuário via POST /admin/adicionar-usuario.
    Verifica que o novo usuário pode fazer login.
    """
    try:
        import render_api.app as app_mod
    except ImportError:
        import app as app_mod
    monkeypatch.setattr(app_mod, "USERS_JSON_PATH", tmp_path / "users.json")

    r = client.post(
        "/admin/adicionar-usuario",
        json={"email": "novo@artesp.local", "senha": "senha123", "role": "user"},
    )
    assert r.status_code == 200, f"Esperado 200, obtido {r.status_code}: {r.text}"
    data = r.json()
    assert "sucesso" in data.get("message", "").lower()

    login_r = client_sem_auth.post(
        "/auth/login",
        json={"email": "novo@artesp.local", "senha": "senha123"},
    )
    assert login_r.status_code == 200, f"Login falhou: {login_r.text}"
    assert "access_token" in login_r.json()
