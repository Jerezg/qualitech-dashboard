@echo off
title Qualitech - Dashboard de Performance Operacional
cd /d "%~dp0"
echo.
echo ============================================================
echo   Qualitech - Dashboard de Performance Operacional
echo ============================================================
echo.

rem --------------------------------------------------------------
rem O comando "streamlit" sozinho so funciona se a pasta "Scripts"
rem do Python estiver na PATH do Windows - o que nem sempre acontece.
rem Para evitar esse problema, chamamos sempre "python -m streamlit"
rem (ou "py -m streamlit"), que funciona mesmo sem isso, desde que o
rem Python em si esteja instalado.
rem --------------------------------------------------------------

set PYCMD=

where python >nul 2>nul
if %errorlevel%==0 (
    set PYCMD=python
) else (
    where py >nul 2>nul
    if %errorlevel%==0 (
        set PYCMD=py
    )
)

if "%PYCMD%"=="" (
    echo [ERRO] Nao encontrei o Python instalado nesta maquina.
    echo.
    echo Instale o Python em https://www.python.org/downloads/
    echo IMPORTANTE: na tela de instalacao, marque a opcao
    echo "Add python.exe to PATH" antes de clicar em Install.
    echo.
    echo Depois de instalar, feche esta janela e de duplo clique
    echo neste arquivo novamente.
    echo.
    pause
    exit /b 1
)

echo Usando: %PYCMD%
%PYCMD% --version

rem Detecta o Python da Microsoft Store, que instala tudo numa pasta
rem "AppData\Local\Packages\...\LocalCache\..." MUITO longa. O Windows
rem tem um limite classico de 260 caracteres pra caminhos de arquivo, e
rem alguns pacotes Python (com nomes de arquivo internos longos) esbarram
rem nesse limite so nessa versao do Python - da erro tipo "OSError:
rem [Errno 2] No such file or directory" na hora de instalar dependencias.
for /f "delims=" %%P in ('%PYCMD% -c "import sys; print(sys.executable)" 2^>nul') do set PYEXE=%%P
echo %PYEXE% | findstr /i "WindowsApps Packages\\PythonSoftwareFoundation" >nul
if %errorlevel%==0 (
    echo.
    echo [AVISO] Este e o Python da Microsoft Store ^(%PYEXE%^).
    echo Essa versao e conhecida por dar erro ao instalar pacotes, por
    echo causa de um limite de tamanho de caminho de arquivo do Windows.
    echo Se a instalacao abaixo falhar, a solucao mais confiavel e:
    echo   1. Desinstalar o "Python 3.x" da Microsoft Store ^(Configuracoes
    echo      -^> Aplicativos -^> procurar "Python" -^> Desinstalar^)
    echo   2. Instalar o Python "de verdade" em
    echo      https://www.python.org/downloads/ marcando a opcao
    echo      "Add python.exe to PATH" na instalacao
    echo   3. Fechar esta janela e dar duplo clique aqui de novo
    echo.
)

echo.
echo Conferindo/instalando as dependencias (so demora na primeira vez)...
%PYCMD% -m pip install -r requirements.txt --quiet --disable-pip-version-check
if not %errorlevel%==0 (
    echo Tentando de outro jeito ^(instalacao "--user"^)...
    %PYCMD% -m pip install --user -r requirements.txt --quiet --disable-pip-version-check
)
if not %errorlevel%==0 (
    echo.
    echo [ERRO] Nao consegui instalar as dependencias automaticamente.
    echo.
    echo Se apareceu um erro do tipo "OSError: [Errno 2] No such file or
    echo directory" com um caminho MUITO longo cheio de "AppData\Local\
    echo Packages\...", o problema e o Python da Microsoft Store ^(veja o
    echo aviso acima^) - a solucao e trocar pelo Python oficial do site
    echo python.org ^(link e passo a passo no aviso acima^).
    echo.
    echo Caso contrario, tente rodar manualmente, num terminal, dentro
    echo desta pasta:
    echo     %PYCMD% -m pip install -r requirements.txt
    echo.
    pause
    exit /b 1
)

echo.
echo Iniciando o App... isso abre uma aba no seu navegador.
echo.
echo IMPORTANTE: NAO FECHE esta janela preta enquanto estiver
echo usando o dashboard. Ela e o "motor" que mantem o App rodando -
echo fechar esta janela (ou o "X" dela) desliga o dashboard, mesmo
echo que a aba do navegador continue aberta.
echo.
echo Para PARAR o dashboard, e so fechar esta janela normalmente
echo (ou apertar Ctrl+C).
echo.
echo ------------------------------------------------------------
echo.

%PYCMD% -m streamlit run app.py

echo.
echo O dashboard foi encerrado. Pode fechar esta janela.
pause
