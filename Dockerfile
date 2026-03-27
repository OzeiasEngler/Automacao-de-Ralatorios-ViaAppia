FROM python:3.11-slim

# Impede que o Python gere arquivos .pyc e permite logs em tempo real
ENV PYTHONDONTWRITEBYTECODE 1
ENV PYTHONUNBUFFERED 1
# Em imagem Docker é normal instalar como root; evita aviso ruidoso do pip no build
ENV PIP_ROOT_USER_ACTION=ignore

# Dependências de sistema necessárias para geopandas/fiona e pymupdf
RUN apt-get update && apt-get install -y --no-install-recommends \
        libgdal-dev \
        gdal-bin \
        libgeos-dev \
        libproj-dev \
        gcc \
        g++ \
    && rm -rf /var/lib/apt/lists/*

# Pasta da aplicação (raiz do repositório)
WORKDIR /app

# Instala dependências Python
COPY requirements.txt .
RUN pip install --no-cache-dir --upgrade pip && \
    pip install --no-cache-dir -r requirements.txt

# Copia o código do projeto
COPY . .

# Não declarar VOLUME /data: o Render gerencia o disco persistente pelo painel.
# Declarar VOLUME no Dockerfile pode causar shadowing e esconder os dados reais.

# Expõe a porta que o Render exige
EXPOSE 10000

# App FastAPI está em render_api.app (módulo render_api, variável app)
CMD ["uvicorn", "render_api.app:app", "--host", "0.0.0.0", "--port", "10000"]