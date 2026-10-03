# Empacota o leitor de etiquetas para os computadores do setor (sem Python instalado).
#
# Gera dist\leitor_etiquetas\ (PyInstaller --onedir) com um versao.txt. O estoque de PVB
# (planilha SGQ) compara esse versao.txt com o da copia local de cada maquina e so recopia
# a pasta quando ele muda -- entao TODA publicacao precisa de um versao.txt novo.
#
# Uso:  powershell -ExecutionPolicy Bypass -File empacotar.ps1

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

$python = Join-Path $PSScriptRoot ".venv\Scripts\python.exe"
if (-not (Test-Path $python)) { throw "Ambiente .venv nao encontrado em $PSScriptRoot" }

& $python -m PyInstaller --noconfirm --onedir --name leitor_etiquetas `
    --collect-data rapidocr --distpath dist --workpath build --specpath build `
    leitor_etiquetas.py
if ($LASTEXITCODE -ne 0) { throw "PyInstaller falhou (codigo $LASTEXITCODE)" }

$pasta = Join-Path $PSScriptRoot "dist\leitor_etiquetas"
$arquivoVersao = Join-Path $pasta "versao.txt"

# AAAA-MM-DD.N -- N sobe quando ja houve publicacao no mesmo dia
$hoje = Get-Date -Format "yyyy-MM-dd"
$n = 1
$anterior = Join-Path $PSScriptRoot "versao_publicada.txt"
if (Test-Path $anterior) {
    $ultima = (Get-Content $anterior -TotalCount 1).Trim()
    if ($ultima -match "^$hoje\.(\d+)$") { $n = [int]$Matches[1] + 1 }
}
$versao = "$hoje.$n"
# ASCII sem BOM: o VBA le a primeira linha e compara texto
[System.IO.File]::WriteAllText($arquivoVersao, $versao)
[System.IO.File]::WriteAllText($anterior, $versao)

$arquivos = Get-ChildItem $pasta -Recurse -File
$mb = [math]::Round(($arquivos | Measure-Object Length -Sum).Sum / 1MB, 1)
Write-Host ""
Write-Host "Pacote: $pasta"
Write-Host "Versao: $versao   Tamanho: $mb MB   Arquivos: $($arquivos.Count)"
Write-Host "Publicar: copiar a pasta inteira para '<unidade da rede>\leitor etiquetas\'."
