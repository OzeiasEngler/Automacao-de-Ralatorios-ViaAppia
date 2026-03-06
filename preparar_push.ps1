# Prepara o repositório para push: limpa staging e re-adiciona respeitando .gitignore
# Uso: .\preparar_push.ps1

Set-Location $PSScriptRoot

Write-Host "=== Resetando staging (mantém alteracoes em disco) ===" -ForegroundColor Cyan
git reset HEAD

Write-Host "`n=== Re-adicionando arquivos (respeita .gitignore atualizado) ===" -ForegroundColor Cyan
git add .

Write-Host "`n=== Status atual ===" -ForegroundColor Cyan
git status --short

Write-Host "`nPronto. Revisar acima e commitar: git commit -m 'mensagem'" -ForegroundColor Green
Write-Host "Depois: git push origin main" -ForegroundColor Green
