# criar_atalho_area_trabalho.ps1
# ================================
# Cria um atalho "Gerador COAF" na Area de Trabalho do usuario atual,
# apontando para o iniciar_gerador_coaf_oculto.vbs desta MESMA pasta.
#
# Funciona em qualquer computador/usuario sem precisar editar nada: o
# script se localiza sozinho (via $PSScriptRoot), entao o mesmo arquivo
# serve tanto pra esta maquina quanto para as máquinas dos demais analistas, desde
# que a pasta COAF inteira (com este script dentro) seja copiada para la.
#
# COMO RODAR (uma vez em cada computador):
#   1) Botao direito neste arquivo -> "Executar com o PowerShell"
#   OU
#   2) Abrir o PowerShell nesta pasta e rodar:
#      powershell -ExecutionPolicy Bypass -File criar_atalho_area_trabalho.ps1

$ErrorActionPreference = "Stop"

$pastaProjeto = $PSScriptRoot
$vbs = Join-Path $pastaProjeto "iniciar_gerador_coaf_oculto.vbs"

if (-not (Test-Path $vbs)) {
    Write-Host "ERRO: nao encontrei iniciar_gerador_coaf_oculto.vbs em $pastaProjeto"
    Write-Host "Rode este script de dentro da pasta do projeto (COAF)."
    exit 1
}

$areaTrabalho = [Environment]::GetFolderPath("Desktop")
$atalho = Join-Path $areaTrabalho "Gerador COAF.lnk"

$shell = New-Object -ComObject WScript.Shell
$link = $shell.CreateShortcut($atalho)
$link.TargetPath = $vbs
$link.WorkingDirectory = $pastaProjeto
$link.IconLocation = "shell32.dll,220"
$link.Description = "Gerador de Reportes COAF - PLD/FT"
$link.Save()

Write-Host "Atalho criado com sucesso em:"
Write-Host $atalho
Write-Host ""
Write-Host "De um duplo-clique nele na Area de Trabalho para abrir o Gerador COAF."
