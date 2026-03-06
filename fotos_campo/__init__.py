"""
fotos_campo – Módulo de processamento de fotos de campo (inspeção rodoviária).

Replica exatamente o que as Macros Panelas (Coord Renomear) faziam, de forma mais
ágil e consistente: sem depender de Excel/VBA, execução em lote, uso em servidor
e CLI, e mesmo fluxo (listar → coordenadas → copiar/renomear → relatório 2 lados → Kcor).

Fluxo (módulos expostos pelo router):
  1. listar       → XLSX com metadados + GPS EXIF das fotos
  2. coordenadas  → preenche Rodovia/km/Sentido por proximidade de coordenada
  3. renomear     → copia e renomeia arquivos (6 padrões)

Compatível com Python 3.9+.
Sem dependência de Tkinter — pode ser usado em contexto de servidor (web/CLI).
"""
__version__ = "1.0.0"
