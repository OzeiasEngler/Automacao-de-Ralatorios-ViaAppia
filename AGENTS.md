# Orientação para agentes e contribuidores

## Lei do projeto: comentários

**Comentários e docstrings só quando forem profissionais e estritamente necessários** (porquê de negócio, workaround, contrato de API). Não documentar “o que o código já diz”. Evitar blocos longos — isso vira retrabalho para enxugar depois.

Detalhe: `.cursor/rules/comentarios-codigo-limpo.mdc` (`alwaysApply: true`).

## Lei do projeto: testes locais primeiro

Toda alteração **passa por testes locais** antes de commit ou de dar a tarefa por fechada. Na raiz: `python -m pytest` ou `test_local.bat`. O que não estiver na suite pytest deve ser validado manualmente e isso deve ficar explícito.

Detalhe: `.cursor/rules/testes-locais-antes.mdc` (`alwaysApply: true`).
