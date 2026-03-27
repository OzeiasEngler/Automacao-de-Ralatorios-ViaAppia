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

# Copia o código do projeto (templates NC/Kartado: nc_artesp/assets/templates/ e fotos_campo/assets/templates/).
# No Render, ative Git LFS para os .xlsx serem ficheiros reais no build.
COPY . .

# Diagnóstico (não falha o build): M04 usa Acumulado.xlsx ou modelo Kcor-Kria em templates/ — ficheiros <2 KiB costumam ser ponteiro LFS.
RUN python -c "import sys; from pathlib import Path; p=Path('nc_artesp/assets/templates'); xs=(([f for f in p.glob('_Planilha Modelo Kcor-Kria*') if f.is_file()] + ([p/'Acumulado.xlsx'] if (p/'Acumulado.xlsx').is_file() else [])) if p.is_dir() else []); mx=max((f.stat().st_size for f in xs), default=0); print('M04 templates (build):', [(f.name, f.stat().st_size) for f in xs] or 'nenhum', file=sys.stderr); (mx<=2048) and print('WARNING: Acumulado.xlsx ou Kcor-Kria ausente ou <2 KiB (LFS?).', file=sys.stderr)"

# Não declarar VOLUME /data: o Render gerencia o disco persistente pelo painel.
# Declarar VOLUME no Dockerfile pode causar shadowing e esconder os dados reais.

# Expõe a porta que o Render exige
EXPOSE 10000

# App FastAPI está em render_api.app (módulo render_api, variável app)
CMD ["uvicorn", "render_api.app:app", "--host", "0.0.0.0", "--port", "10000"]