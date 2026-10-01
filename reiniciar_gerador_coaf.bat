@echo off
cd /d "%~dp0"

echo Encerrando instancias antigas do Gerador COAF (Streamlit)...
powershell -NoProfile -Command "Get-CimInstance Win32_Process | Where-Object { $_.CommandLine -match 'streamlit' -and $_.CommandLine -match 'gerador_coaf' } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force }"

if exist "__pycache__" rmdir /s /q "__pycache__"

echo Iniciando versao atual...
call "iniciar_gerador_coaf.bat"
