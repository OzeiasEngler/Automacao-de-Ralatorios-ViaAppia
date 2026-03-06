# GeradorARTESP.exe e GeoJSON reduzido

## Origem do executável

O **GeradorARTESP.exe** (ex.: em `OneDrive\...\Gerador geojson\dist\dist\GeradorARTESP.exe`) é gerado por **PyInstaller** a partir deste repositório:

- **Entrada:** `gerador_artesp_gui.py`
- **Incluído nos dados:** `gerador_artesp_core.py` + pasta `assets` (+ pyarmor se usado)
- **Arquivo de build:** `GeradorARTESP.spec`

Ou seja: o .exe usa o **core** que está neste repo no momento do build.

---

## Por que o GeoJSON pode sair com 2,6 MB

Se o .exe foi compilado com uma **versão antiga** do `gerador_artesp_core.py` (sem redução na gravação), ao salvar o GeoJSON:

- Não há redução de pontos (step)
- Coordenadas com mais decimais

Resultado: arquivos maiores (~2,6 MB no seu caso).

---

## O que este repositório tem hoje

No **`gerador_artesp_core.py`** atual a função `salvar_geojson` já:

1. **Reduz pontos** (step): mantém 1º, último e cada 2º ponto (step=2 por padrão).
2. **Usa 4 decimais** nas coordenadas (lon/lat).
3. **Respeita variáveis de ambiente** (útil mesmo no .exe):
   - `ARTESP_GEOJSON_SIMPLIFY_STEP` (padrão 2; use 1 para não reduzir pontos).
   - `ARTESP_GEOJSON_DECIMAIS` (padrão 4; entre 3 e 6).

Com isso, o mesmo conjunto de dados tende a gerar GeoJSON bem menor (na faixa de 40–50% de redução).

---

## O que fazer para o .exe gerar GeoJSON menor

### 1. Recompilar o .exe a partir deste repo

Assim o executável passa a usar o **core atual** (com redução):

1. Abra o projeto em: `c:\GeradorARTESP`
2. Garanta que `gerador_artesp_core.py` está atualizado (com `_simplificar_coordenadas` e `salvar_geojson` com step/decimais).
3. Gere o .exe com PyInstaller, por exemplo:
   ```bat
   pyinstaller GeradorARTESP.spec
   ```
4. O .exe sairá em `dist\GeradorARTESP.exe` (ou em `dist\dist\` conforme seu .spec).
5. **Substitua** o executável na pasta do OneDrive pelo novo .exe.

A partir daí, o .exe do OneDrive passará a gerar GeoJSON já reduzido (step=2, 4 decimais).

---

### 2. Ajustar redução sem recompilar (opcional)

Quem usa o .exe pode controlar a redução por **variáveis de ambiente** antes de abrir o programa:

- **Menos redução** (arquivo um pouco maior, mais fiel ao traçado):
  - `ARTESP_GEOJSON_SIMPLIFY_STEP=1` (mantém todos os pontos).
- **Mais redução** (arquivo menor):
  - `ARTESP_GEOJSON_SIMPLIFY_STEP=3` (mantém cada 3º ponto).
  - `ARTESP_GEOJSON_DECIMAIS=3` (3 decimais nas coordenadas).

No Windows, dá para definir no atalho que abre o .exe, por exemplo:

```bat
set ARTESP_GEOJSON_SIMPLIFY_STEP=2
set ARTESP_GEOJSON_DECIMAIS=4
"c:\Users\oseia\OneDrive - VIA APPIA CONCESSOES S.A\...\dist\dist\GeradorARTESP.exe"
```

Ou criar um `.bat` na mesma pasta do .exe com essas linhas e abrir o Gerador por esse .bat.

---

## Resumo

| Objetivo | Ação |
|----------|------|
| GeoJSON menor pelo .exe | Recompilar o .exe a partir de **c:\GeradorARTESP** (core atual) e trocar o .exe no OneDrive. |
| Ajustar grau de redução | Usar `ARTESP_GEOJSON_SIMPLIFY_STEP` e `ARTESP_GEOJSON_DECIMAIS` ao rodar o .exe (atalho ou .bat). |

O “render”/relatório que sai do **GeradorARTESP.exe** nessa pasta do OneDrive passa a usar a lógica de redução assim que você usar um .exe gerado a partir deste repositório atualizado.
