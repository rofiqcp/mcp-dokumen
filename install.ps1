$ErrorActionPreference = "Stop"

function Write-Info {
    param([string]$Message)
    Write-Host "[INFO] $Message"
}

function Write-Warn {
    param([string]$Message)
    Write-Warning $Message
}

$RepoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$pythonCandidates = @("python3.10", "python3", "python")
$pythonCmd = $null

foreach ($candidate in $pythonCandidates) {
    try {
        $ver = & $candidate -c "import sys; print('.'.join(map(str, sys.version_info[:3])))" 2>$null
        if (-not $ver) { continue }
        $verObj = [Version]$ver
        if ($verObj.Major -gt 3 -or ($verObj.Major -eq 3 -and $verObj.Minor -ge 10)) {
            $pythonCmd = $candidate
            break
        }
    } catch {
        continue
    }
}

if (-not $pythonCmd) {
    Write-Warn "Python 3.10+ tidak ditemukan. Instal Python 3.10.x terlebih dahulu (lihat tautan di README)."
    exit 1
}

Write-Info "Menggunakan Python: $pythonCmd"

$venvPath = Join-Path $RepoRoot ".venv"
if (-not (Test-Path $venvPath)) {
    Write-Info "Membuat virtual environment di $venvPath"
    & $pythonCmd -m venv $venvPath
}

$pythonExe = Join-Path $venvPath "Scripts\python.exe"

Write-Info "Memperbarui pip"
& $pythonExe -m pip install --upgrade pip

Write-Info "Meng-install paket dalam mode editable"
& $pythonExe -m pip install -e $RepoRoot

# Path VS Code MCP config: C:\Users\$env:USERNAME\AppData\Roaming\Code\User\mcp.json
$configPath = "C:\Users\$env:USERNAME\AppData\Roaming\Code\User\mcp.json"

$configDir = Split-Path -Parent $configPath
if (-not (Test-Path $configDir)) {
    New-Item -ItemType Directory -Force -Path $configDir | Out-Null
}

if (Test-Path $configPath) {
    Copy-Item $configPath "$configPath.bak" -Force
}

$config = @{}
if (Test-Path $configPath -PathType Leaf -and (Get-Item $configPath).Length -gt 0) {
    try {
        $config = Get-Content $configPath -Raw | ConvertFrom-Json
    } catch {
        Write-Warn "Gagal membaca $configPath, menggunakan konfigurasi baru."
        $config = @{}
    }
}

if (-not ($config.PSObject.Properties.Name -contains "servers")) {
    $config | Add-Member -MemberType NoteProperty -Name "servers" -Value (@{})
}

$config.servers.mcpdocx = @{
    type    = "stdio"
    command = $pythonExe
    args    = @("-m", "mcpdocx")
    cwd     = $RepoRoot
    env     = @{
        PYTHONPATH = $RepoRoot
    }
}

if (-not ($config.PSObject.Properties.Name -contains "inputs")) {
    $config | Add-Member -MemberType NoteProperty -Name "inputs" -Value @()
}

# Hapus mcpServers jika ada (format lama tidak menggunakannya)
if ($config.PSObject.Properties.Name -contains "mcpServers") {
    $config.PSObject.Properties.Remove("mcpServers")
}

$config | ConvertTo-Json -Depth 6 | Set-Content -Path $configPath -Encoding UTF8

Write-Info "Konfigurasi MCP disimpan ke $configPath"
Write-Info "Instalasi selesai. Aktivasi venv: .\\.venv\\Scripts\\Activate.ps1"
