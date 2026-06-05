git add .
$msg = Read-Host "Mensagem do commit"
git commit -m $msg
git pull origin main
git push origin main
Write-Host "Push concluido!" -ForegroundColor Green