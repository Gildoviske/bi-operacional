@echo off
cd /d "%~dp0"
echo Lendo os relatorios dos Correios...
python atualizar_rastreios.py
if errorlevel 1 (
    echo ERRO ao ler os rastreios. Verifique a pasta CONTROLE COMPRAS\rastreios e se a planilha de transferencias nao esta aberta.
    pause
    exit /b 1
)
echo.
echo Gerando pagina a partir das planilhas...
python gerar_pagina.py
if errorlevel 1 (
    echo ERRO ao gerar a pagina. Verifique se as planilhas nao estao abertas com erro.
    pause
    exit /b 1
)
echo Publicando no GitHub...
git add index.html
git commit -m "Atualizacao diaria"
git push
echo.
echo Pronto! Pagina publicada em: https://gildoviske.github.io/bi-operacional/
pause
