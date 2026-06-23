@echo off
cd /d "%~dp0"
set "PYTHON_EXE=C:\Users\avail\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
set "URL=http://127.0.0.1:5500/"

echo.
echo Iniciando Painel Operacional local...
echo Pasta: %CD%
echo Endereco: %URL%
echo.
echo Mantenha esta janela aberta enquanto estiver testando.
echo Para parar o servidor, feche esta janela.
echo.

start "" "%URL%"

if exist "%PYTHON_EXE%" (
  "%PYTHON_EXE%" -m http.server 5500 --bind 127.0.0.1
) else (
  where py >nul 2>nul
  if errorlevel 1 (
    echo.
    echo ERRO: Nao encontrei Python neste computador.
    echo Instale Python ou abra este projeto pelo Live Server do VS Code.
    echo.
    pause
    exit /b 1
  )
  py -m http.server 5500 --bind 127.0.0.1
)

echo.
echo O servidor foi encerrado.
pause
