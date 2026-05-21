@echo off
title Ober TI - Gestao de Usuarios

:: Solicitar Permissoes de Administrador Automaticamente (Auto-Elevate UAC)
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if '%errorlevel%' NEQ '0' (
    goto UACPrompt
) else ( goto gotAdmin )

:UACPrompt
    echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
    echo UAC.ShellExecute "cmd.exe", "/c ""%~s0""", "", "runas", 1 >> "%temp%\getadmin.vbs"
    "%temp%\getadmin.vbs"
    exit /B

:gotAdmin
    if exist "%temp%\getadmin.vbs" ( del "%temp%\getadmin.vbs" )
    pushd "%CD%"
    CD /D "%~dp0"

:: Desbloqueia o script caso tenha sido bloqueado pelo Windows Defender/SmartScreen
powershell -NoProfile -Command "Unblock-File -Path '%~dp0ober_gestao_usuarios.ps1'"

:: Executa o script ignorando a politica de restricao local
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0ober_gestao_usuarios.ps1"

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo === ERRO AO EXECUTAR O SCRIPT ===
    echo Codigo de erro: %ERRORLEVEL%
    echo.
    pause
)