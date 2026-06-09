git add .
$msg = Read-Host "Mensagem do commit"
git commit -m $msg
git pull origin main
if ($LASTEXITCODE -eq 0) {
    git push origin main
    if ($LASTEXITCODE -eq 0) {
        Write-Host "Push concluido com sucesso!" -ForegroundColor Green
    } else {
        Write-Host "Erro no push!" -ForegroundColor Red
    }
} else {
    Write-Host "Erro no pull! Resolva os conflitos antes de continuar." -ForegroundColor Red
}
pause