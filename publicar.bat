@echo off
cd /d "%~dp0"
echo Gerando pagina a partir das planilhas...
python gerar_pagina.py
if errorlevel 1 (
    echo ERRO ao gerar a pagina. Verifique se as planilhas nao estao abertas com erro.
    pause
    exit /b 1
)
echo Publicando no GitHub...
git add index.html historico_inventario.json historico_inventario_acessorios.json
git commit -m "Atualizacao diaria"
git push
echo.
echo Pronto! Pagina publicada em: https://gildoviske.github.io/bi-operacional/
pause
