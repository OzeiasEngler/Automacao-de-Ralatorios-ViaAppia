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

# Copia o código do projeto (inclui templates Kartado: nc_artesp/assets/templates/,
# fotos_campo/assets/Template/ — no Render, ative Git LFS no serviço para os .xlsx serem ficheiros reais no build).
COPY . .

# Acumulado M04 / calendário M06 exigem «_Planilha Modelo Kcor-Kria» real no disco (não só ponteiro LFS ~130 B).
RUN python -c "from pathlib import Path; p=Path('nc_artesp/assets'); xs=[f for f in p.glob('_Planilha Modelo Kcor-Kria*') if f.is_file()]; mx=max((f.stat().st_size for f in xs), default=0); assert mx>2048, ('Kcor-Kria template em falta ou ponteiro LFS no build — ative Git LFS no Render e redeploy. Ficheiros: %r' % ([(f.name, f.stat().st_size) for f in xs],))"

# Não declarar VOLUME /data: o Render gerencia o disco persistente pelo painel.
# Declarar VOLUME no Dockerfile pode causar shadowing e esconder os dados reais.

# Expõe a porta que o Render exige
EXPOSE 10000

# App FastAPI está em render_api.app (módulo render_api, variável app)
CMD ["uvicorn", "render_api.app:app", "--host", "0.0.0.0", "--port", "10000"]