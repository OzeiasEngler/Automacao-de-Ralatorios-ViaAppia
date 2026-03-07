# Deploy da API na Locaweb (VPS + Docker)

A Locaweb **não suporta Python em hospedagem compartilhada**. Para rodar esta API é necessário **VPS** ou **Cloud Server Pro**. O guia abaixo usa VPS com Docker (igual ao que já existe no projeto).

## Pré-requisitos

- Conta Locaweb com **VPS** ou **Cloud Server Pro**
- Acesso SSH ao servidor (chave ou senha)
- Domínio ou IP para acessar a API (opcional: só o IP já serve)

## 1. Criar e acessar o VPS

1. No [painel Locaweb](https://painel.locaweb.com.br), crie um **VPS** (Ubuntu 22.04 é uma boa opção).
2. Anote o **IP** e as credenciais SSH.
3. Se for usar domínio, aponte um subdomínio (ex.: `api.seudominio.com.br`) para esse IP no DNS.
4. Conecte por SSH:
   ```bash
   ssh root@SEU_IP
   ```
   (ou o usuário que você configurou)

## 2. Instalar Docker no VPS

Documentação oficial: [Instalar Docker – Locaweb](https://www.locaweb.com.br/ajuda/wiki/instalar-docker/).

Resumo (Ubuntu):

```bash
sudo apt update
sudo apt install -y apt-transport-https ca-certificates curl software-properties-common
curl -fsSL https://download.docker.com/linux/ubuntu/gpg | sudo gpg --dearmor -o /usr/share/keyrings/docker-archive-keyring.gpg
echo "deb [arch=$(dpkg --print-architecture) signed-by=/usr/share/keyrings/docker-archive-keyring.gpg] https://download.docker.com/linux/ubuntu $(lsb_release -cs) stable" | sudo tee /etc/apt/sources.list.d/docker.list > /dev/null
sudo apt update
sudo apt install -y docker-ce docker-ce-cli containerd.io
sudo systemctl enable docker && sudo systemctl start docker
```

(Opcional) Docker Compose:

```bash
sudo apt install -y docker-compose-plugin
```

## 3. Enviar o projeto para o VPS

**Opção A – Git (recomendado)**

```bash
sudo apt install -y git
cd /opt   # ou outro diretório de sua preferência
sudo git clone https://github.com/OzeiasEngler/Automacao-de-Ralatorios-ViaAppia.git api-artesp
cd api-artesp
```

**Opção B – SCP/SFTP**

No seu PC (PowerShell ou terminal):

```powershell
scp -r C:\GeradorARTESP\* usuario@SEU_IP:/opt/api-artesp/
```

(Substitua `usuario` e `SEU_IP`; crie antes a pasta `/opt/api-artesp` no servidor.)

## 4. Variáveis de ambiente

A API usa variáveis para dados persistentes e segurança. Crie um arquivo `.env` **no servidor** (fora do repositório), por exemplo em `/opt/api-artesp/.env`:

```env
# Dados persistentes (usuários, métricas, arquivos gerados)
ARTESP_DATA_DIR=/data
ARTESP_OUTPUT_DIR=/data/outputs

# Segurança: chave para JWT (gere um valor forte, ex.: openssl rand -hex 32)
JWT_SECRET=sua_chave_secreta_aqui_32_caracteres_ou_mais
```

Crie as pastas que a API espera:

```bash
sudo mkdir -p /data/outputs
sudo chown -R 1000:1000 /data   # ou o UID que o container usar
```

## 5. Rodar com Docker

A API escuta na **porta 10000** (definida no `Dockerfile`).

**Usando apenas Docker:**

```bash
cd /opt/api-artesp

# Build da imagem
sudo docker build -t api-artesp .

# Container com volume para dados e .env
sudo docker run -d \
  --name api-artesp \
  --restart unless-stopped \
  -p 10000:10000 \
  -v /data:/data \
  -e ARTESP_DATA_DIR=/data \
  -e ARTESP_OUTPUT_DIR=/data/outputs \
  -e JWT_SECRET=sua_chave_secreta_aqui \
  api-artesp
```

**Usando Docker Compose** (se criou o arquivo `docker-compose.locaweb.yml`):

```bash
cd /opt/api-artesp
sudo docker compose -f docker-compose.locaweb.yml up -d
```

## 6. Testar

- Pelo IP: `http://SEU_IP:10000`
- Se configurou domínio: `http://api.seudominio.com.br:10000`
- Health/docs: `http://SEU_IP:10000/docs`

## 7. Deixar na porta 80 (opcional)

Para acessar sem `:10000`, use um proxy reverso. Exemplo com **nginx**:

```bash
sudo apt install -y nginx
```

Crie `/etc/nginx/sites-available/api-artesp`:

```nginx
server {
    listen 80;
    server_name api.seudominio.com.br;   # ou _ para aceitar qualquer host
    location / {
        proxy_pass http://127.0.0.1:10000;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }
}
```

Ative e recarregue:

```bash
sudo ln -s /etc/nginx/sites-available/api-artesp /etc/nginx/sites-enabled/
sudo nginx -t && sudo systemctl reload nginx
```

Assim a API fica em `http://api.seudominio.com.br` (porta 80).

## 8. HTTPS (recomendado)

Com nginx, use **Certbot** (Let’s Encrypt):

```bash
sudo apt install -y certbot python3-certbot-nginx
sudo certbot --nginx -d api.seudominio.com.br
```

## 9. CORS — usar a API/GeoJSON em um site em outro domínio

Se o **mapa (Leaflet, Google Maps, etc.)** ou outro front está em um site hospedado em outro domínio (ex.: site na Locaweb), o navegador só permite a requisição se a API liberar a origem (CORS).

- **Opção recomendada:** defina a variável **`ARTESP_CORS_ORIGINS`** com a(s) origem(ns) do site, separadas por vírgula, **sem espaço** após a vírgula:
  ```env
  ARTESP_CORS_ORIGINS=https://www.seudominio.com.br,https://seudominio.com.br,https://meusite.locaweb.com.br
  ```
  Use exatamente a URL que aparece na barra do navegador (com ou sem `www`, `http` ou `https`).

- **Para testes / aceitar qualquer site:** use o valor `*`:
  ```env
  ARTESP_CORS_ORIGINS=*
  ```
  A API passará a refletir o header `Origin` da requisição (qualquer domínio pode chamar a API). Em produção prefira lista explícita.

No **Render**: Dashboard do serviço → Environment → Add variable `ARTESP_CORS_ORIGINS` com um dos valores acima e faça redeploy.

## Resumo rápido

| Item              | Valor                    |
|-------------------|--------------------------|
| Tipo de hospedagem | VPS ou Cloud Server Pro |
| Porta da API      | 10000                    |
| Dados persistentes | `/data` (volume Docker) |
| Variáveis importantes | `ARTESP_DATA_DIR`, `ARTESP_OUTPUT_DIR`, `JWT_SECRET`, `ARTESP_CORS_ORIGINS` |

Documentação Locaweb: [Docker no VPS](https://www.locaweb.com.br/ajuda/wiki/instalar-docker/), [API de Servidores](https://developer.locaweb.com.br/documentacoes/apiservidores).
