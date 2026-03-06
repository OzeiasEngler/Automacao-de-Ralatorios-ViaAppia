# Avaliação: código que gera GeoJSON ~2,6 MB

## O que define o tamanho do arquivo

1. **Número de features**  
   Uma feature por linha do Excel → muitas linhas = muitas features = mais JSON (properties + geometria).

2. **Pontos por geometria**  
   No código colado:
   - `MAX_PONTOS_POR_LINHA = 400` → cada LineString pode ter até 400 pontos.
   - `simplificar()` (Douglas–Peucker 10 m) e `filtrar_espaco_minimo(50 m)` já reduzem na **extração**, mas cada trecho ainda pode ter dezenas/centenas de pontos.

3. **Formato na gravação (`salvar_geojson`)**  
   No código colado, ao salvar:
   - Só há **sanitize**: arredondar floats (`round(o, 6)` para |x|>20, senão `round(o, 3)`).
   - **Não há** redução de pontos na hora de salvar (nenhum “step” nem limite de decimais configurável).
   - `separators=(',',':')` já é compacto (sem espaços).

4. **Coordenadas**  
   Longitude/latitude com 6 decimais = ~15–20 caracteres por coordenada; muitas features × muitos pontos = tamanho alto.

**Conclusão:** O 2,6 MB vêm de muitas features, com muitas coordenadas por linha (até 400 pontos) e 6 decimais, sem nova redução na gravação.

---

## O que o `gerador_artesp_core.py` do repositório já faz

No **core** do repo a função `salvar_geojson` já:

- Aplica **redução por step** (`_simplificar_coordenadas`): mantém 1º, último e cada `step`-ésimo ponto (ex.: step=2 → ~metade dos pontos).
- Usa **decimais configuráveis** (padrão 4 para lon/lat), reduzindo caracteres por número.
- Lê opcionalmente:
  - `ARTESP_GEOJSON_SIMPLIFY_STEP` (padrão 2),
  - `ARTESP_GEOJSON_DECIMAIS` (padrão 4).

Quem gera o GeoJSON **usando o core** (ex.: `render_api` ou script que chama `gerador_artesp_core.salvar_geojson`) já se beneficia dessa redução.

---

## Se você usa o script “colado” (standalone) em vez do core

Se o GeoJSON de 2,6 MB é gerado por esse script que tem apenas:

```python
def salvar_geojson(caminho, obj):
    def sanitize(o): ...
    json.dump(sanitize(obj), f, ...)
```

então **não há redução na gravação**. Para reduzir:

1. **Substituir `salvar_geojson`** pela lógica do core: step (ex. 2) + decimais (ex. 4) + `_simplificar_coordenadas` na gravação.
2. **Opcional:** reduzir `MAX_PONTOS_POR_LINHA` (ex.: 400 → 200) ou aumentar um pouco a tolerância do Douglas–Peucker para menos pontos já na extração.

---

## Resumo

| Fator                         | Código colado      | Core do repo              |
|------------------------------|--------------------|----------------------------|
| Redução de pontos ao salvar  | Não                | Sim (step, ex. 2)          |
| Decimais das coordenadas     | 6 / 3              | 4 (configurável)           |
| Limite de pontos na extração | 400                | 5000 (redução no save)     |
| Tamanho esperado             | Maior (~2,6 MB)    | Menor com step=2 e 4 dec.  |

Para o mesmo conjunto de dados, usar a `salvar_geojson` do **gerador_artesp_core.py** (com step=2 e 4 decimais) deve reduzir o tamanho do GeoJSON de forma relevante (na ordem de ~40–50%, dependendo do número de pontos por feature).
