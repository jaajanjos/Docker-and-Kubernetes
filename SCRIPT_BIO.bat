@echo off
REM ==========================================================
REM  SCRIPT_BIO.bat
REM  Baixa os arquivos do Projeto BIO do GitHub e copia para
REM  D:\PROJETOS_CODE\BIO  (nao precisa de Git instalado)
REM  Requisito: Windows 10 ou 11 (curl e tar ja vem no Windows)
REM  Uso: de dois cliques no arquivo ou rode no CMD: SCRIPT_BIO.bat
REM ==========================================================
setlocal
chcp 65001 >nul

set "DEST=D:\PROJETOS_CODE\BIO"
set "URL=https://github.com/jaajanjos/Docker-and-Kubernetes/archive/refs/heads/claude/funny-brown-vetnlo.zip"
set "ZIP=%TEMP%\bio.zip"
set "TMPDIR=%TEMP%\bio_extract"
set "SRC=%TMPDIR%\Docker-and-Kubernetes-claude-funny-brown-vetnlo"

echo.
echo ===== PROJETO BIO - ATUALIZACAO =====
echo Destino: %DEST%
echo.

if not exist "D:\" (
    echo [ERRO] Unidade D: nao encontrada.
    goto :fim
)

if not exist "%DEST%" mkdir "%DEST%"

echo [1/4] Baixando arquivos do GitHub...
curl -L -s -f -o "%ZIP%" "%URL%"
if errorlevel 1 (
    echo [ERRO] Falha no download. Verifique sua conexao com a internet.
    goto :fim
)

echo [2/4] Extraindo...
if exist "%TMPDIR%" rmdir /s /q "%TMPDIR%"
mkdir "%TMPDIR%"
tar -xf "%ZIP%" -C "%TMPDIR%"
if errorlevel 1 (
    echo [ERRO] Falha ao extrair o arquivo ZIP.
    goto :limpar
)

echo [3/4] Copiando para %DEST% ...
for %%P in ("Portfolio_Freelancer" "E-COMMERCE EM EXCEL" "PLANO DE ESTUDOS") do (
    robocopy "%SRC%\%%~P" "%DEST%\%%~P" /E /NFL /NDL /NJH /NJS /NP >nul
    echo       - %%~P
)
for %%F in ("PROJETO_BIO.md" "PROJETO_BIO.pdf" "DOCUMENTO_GERAL.docx" "SCRIPT_BIO.bat") do (
    copy /Y "%SRC%\%%~F" "%DEST%\" >nul
    echo       - %%~F
)

:limpar
echo [4/4] Limpando arquivos temporarios...
if exist "%ZIP%" del /q "%ZIP%"
if exist "%TMPDIR%" rmdir /s /q "%TMPDIR%"

echo.
echo ===== CONCLUIDO =====
echo Arquivos em: %DEST%
start "" explorer "%DEST%"

:fim
echo.
pause
endlocal
