# Padrão GeoJSON (referência para Fotos e NC)

O **GeoJSON** (`index.html`) é a referência em produção no Render. Todas as páginas de bloco (Fotos, NC) devem seguir o **mesmo padrão** de layout e log.

## Estrutura da página (GeoJSON)

- **CSS:** `bloco-padrao.css` (design system único).
- **Layout:** `.page-content` → `.grid-2` → card esquerda (upload + opções + botões + result-area) + card direita (Log).
- **Card do log:**
  - `.card-header`: "Log de Processamento" + botão "Limpar" (`id="btnLimparLog"`).
  - `.card-body` → `.log-container` com `id="logBox"`.
  - Estado vazio: um único filho `.log-empty` com `<i class="fas fa-terminal"></i>` e `<span>Nenhuma atividade registrada</span>`.

## Comportamento do log (GeoJSON)

- **logMsg(text, type):** 2 parâmetros. Usa sempre `logBox`. Na primeira mensagem, limpa o container e remove o `.log-empty`. Adiciona linha com `[HH:MM:SS]` + texto, classe `.log-line` + tipo (info, success, error, warning). Faz scroll para o final.
- **Limpar log:** `logBox.innerHTML = '<div class="log-empty"><i class="fas fa-terminal"></i><span>Nenhuma atividade registrada</span></div>';` e reset do flag `logStarted`.
- **Processamento em tempo real:** GeoJSON usa **SSE** (resposta em stream do endpoint `/gerar-relatorio-progresso`). O servidor envia eventos `data: {"type":"progress","status":"..."}` e o cliente chama `logMsg(ev.status, 'info')` para cada um. Por isso o log enche durante a execução.

## Fotos

- Mesma estrutura: um único `logBox`, `logMsg(text, type)` com 2 args, limpar com o mesmo HTML do `.log-empty`.
- Também usa **SSE** (`/fotos/processar-completo-progresso`) para log em tempo real.

## NC

- **Múltiplos painéis** → múltiplos logs (`log-analisar-pdf`, `log-extrair-pdf`, `log-separar`, `log-modelo-foto`, `log-conservacao`, `log-meio-ambiente`, `log-juntar`, `log-inserir-numero`, `log-exportar-calendario`, `log-organizar-imagens`, `log-pipeline`).
- **logMsg(logId, text, type):** 3 parâmetros (id do container). Comportamento igual: remove `.log-empty` se existir, adiciona linha com hora e texto, scroll.
- **limparLog(id):** restaura o mesmo HTML do GeoJSON: `<div class="log-empty"><i class="fas fa-terminal"></i><span>Nenhuma atividade registrada</span></div>`.
- **Processamento:** Endpoints NC hoje retornam só no final (sem SSE). Por isso o log não enche em tempo real como no GeoJSON. Para igualar ao GeoJSON seria necessário o backend enviar eventos de progresso (SSE) nos endpoints NC.

## Ao alterar funções de log no NC

- **logMsg(logId, text, type)** é usada em dezenas de pontos no `nc.html`. Ao mudar assinatura ou comportamento, buscar por `logMsg(` e atualizar todos os chamadores.
- **limparLog(id)** é chamada por todos os botões "Limpar" de cada painel e nos resets. Garantir que o HTML restaurado seja sempre o mesmo do GeoJSON (texto "Nenhuma atividade registrada").
