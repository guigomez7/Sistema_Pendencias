@echo off
chcp 65001 >nul
title Sistema de Envio de Pendencias
cd /d "%~dp0"

echo ============================================================
echo   Sistema de Envio Automatico de Pendencias
echo ============================================================
echo.

REM Verifica se o Python esta instalado e acessivel no PATH
where python >nul 2>nul
if errorlevel 1 (
    echo ERRO: Python nao foi encontrado no PATH deste computador.
    echo Instale o Python em https://www.python.org/downloads/
    echo e marque a opcao "Add Python to PATH" durante a instalacao.
    echo.
    pause
    exit /b 1
)

echo Verificando/instalando dependencias necessarias...
python -m pip install --disable-pip-version-check -q -r requirements.txt
if errorlevel 1 (
    echo.
    echo AVISO: houve um problema ao instalar as dependencias.
    echo Verifique sua conexao com a internet e tente novamente.
    echo.
    pause
    exit /b 1
)

echo Dependencias OK. Iniciando o sistema...
echo.

python enviar_pendencias.py

echo.
echo ============================================================
pause
