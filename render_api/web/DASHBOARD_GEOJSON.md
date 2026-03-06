# Como o dashboard interativo se conecta ao GeoJSON

## 1. Dashboard gerado pelo relatório (recomendado)

Quando você marca **"Gerar dashboard (HTML)"** e gera o relatório:

1. O backend salva o GeoJSON em `outputs/NOME.geojson`.
2. A função `gerar_dashboard_artesp` (core) **lê esse GeoJSON** do disco e **embute os dados dentro do HTML** em uma variável JavaScript:
   ```js
   var geojsonData = { ... };  // GeoJSON inline
   var geoLayer = L.geoJSON(geojsonData, { ... }).addTo(map);
   ```
3. O arquivo `NOME_dashboard.html` é salvo no mesmo diretório (`outputs/`).
4. O link **Dashboard** na tela aponta para `/outputs/NOME_dashboard.html`.

**Conexão:** não há arquivo externo na hora de abrir o dashboard. O GeoJSON já está **dentro** do HTML. Ao abrir o link (ou baixar o arquivo), o mapa funciona sozinho, inclusive offline depois de carregado.

---

## 2. Templates estáticos (L13, L21, L26 em `web/`)

Os arquivos `L13_conservacao_2026_r0_dashboard.html`, `L21_...`, `L26_...` na pasta `web/` são **modelos de layout**. Eles não carregam GeoJSON sozinhos.

Para conectar ao GeoJSON gerado pela API:

- **Opção A:** Use o **dashboard gerado** (item 1): após gerar o relatório, clique no botão **Dashboard** — esse HTML já vem com o GeoJSON embutido.
- **Opção B:** Abra o **visualizador** com parâmetro:  
  `/web/dashboard-viewer?geojson=NOME.geojson`  
  (a página carrega o GeoJSON de `/outputs/NOME.geojson` via fetch; é necessário estar logado).

---

## Resumo

| Origem              | Onde está o GeoJSON?      | Como o dashboard “conecta”        |
|---------------------|---------------------------|-----------------------------------|
| Dashboard gerado    | Dentro do próprio HTML     | Embutido na página (variável JS)  |
| Dashboard viewer    | Arquivo em `/outputs/`     | Fetch para `/outputs/NOME.geojson`|

O fluxo normal é: **Gerar relatório com “Gerar dashboard (HTML)”** → clicar em **Dashboard** → abrir o HTML que já contém o GeoJSON.
