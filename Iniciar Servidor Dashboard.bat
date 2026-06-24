@echo off
title Servidor Dashboard - JET SP

set "PASTA=c:\Users\robson.noberto\Desktop\Controle Kips Motoristas"
set "CLOUDFLARED=C:\Users\robson.noberto\OneDrive - J&T EXPRESS - FILIAL SP\Ana Clara\Lacre saida chegada\cloudflared.exe"

echo Iniciando servidor local na porta 8081...
start "Servidor KPI" cmd /k "cd /d "%PASTA%" && py servidor_dashboard.py"

ping -n 4 127.0.0.1 >nul

echo Iniciando tunel Cloudflare...
echo O link publico aparecera abaixo em alguns segundos...
echo Procure pela linha com: trycloudflare.com
echo.
"%CLOUDFLARED%" tunnel --url http://localhost:8081 2>&1
pause
