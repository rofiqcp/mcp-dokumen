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
$pythonExe = Join-Path $venvPath "Scripts\python.exe"

if (-not (Test-Path $venvPath)) {
    Write-Info "Membuat virtual environment di $venvPath"
    & $pythonCmd -m venv $venvPath
    
    Write-Info "Memperbarui pip"
    & $pythonExe -m pip install --upgrade pip
    
    Write-Info "Meng-install paket dalam mode editable"
    & $pythonExe -m pip install -e $RepoRoot
} else {
    Write-Info "Virtual environment sudah ada di $venvPath"
    
    # Cek apakah paket sudah terinstall
    $installed = & $pythonExe -m pip list 2>$null | Select-String "mcpdocx"
    if (-not $installed) {
        Write-Info "Memperbarui pip"
        & $pythonExe -m pip install --upgrade pip
        
        Write-Info "Meng-install paket dalam mode editable"
        & $pythonExe -m pip install -e $RepoRoot
    } else {
        Write-Info "Paket mcpdocx sudah terinstall, skip instalasi"
    }
}

# Path VS Code MCP config: C:\Users\$env:USERNAME\AppData\Roaming\Code\User\mcp.json
$configPath = "C:\Users\$env:USERNAME\AppData\Roaming\Code\User\mcp.json"

$configDir = Split-Path -Parent $configPath
if (-not (Test-Path $configDir)) {
    New-Item -ItemType Directory -Force -Path $configDir | Out-Null
}

if (Test-Path $configPath) {
    Copy-Item $configPath "$configPath.bak" -Force
}

$config = $null
if ((Test-Path $configPath -PathType Leaf) -and (Get-Item $configPath).Length -gt 0) {
    try {
        $config = Get-Content $configPath -Raw | ConvertFrom-Json
    } catch {
        Write-Warn "Gagal membaca $configPath, menggunakan konfigurasi baru."
        $config = $null
    }
}

# Buat config baru jika tidak ada
if (-not $config) {
    $config = [PSCustomObject]@{
        servers = [PSCustomObject]@{}
        inputs = @()
    }
}

# Pastikan servers dan inputs ada
if (-not ($config.PSObject.Properties.Name -contains "servers")) {
    $config | Add-Member -MemberType NoteProperty -Name "servers" -Value ([PSCustomObject]@{})
}

if (-not ($config.PSObject.Properties.Name -contains "inputs")) {
    $config | Add-Member -MemberType NoteProperty -Name "inputs" -Value @()
}

# Buat konfigurasi mcpdocx
$mcpdocxConfig = [PSCustomObject]@{
    type    = "stdio"
    command = $pythonExe
    args    = @("-m", "mcpdocx")
    cwd     = $RepoRoot
    env     = [PSCustomObject]@{
        PYTHONPATH = $RepoRoot
    }
}

# Tambahkan atau update mcpdocx server
if ($config.servers.PSObject.Properties.Name -contains "mcpdocx") {
    $config.servers.mcpdocx = $mcpdocxConfig
} else {
    $config.servers | Add-Member -MemberType NoteProperty -Name "mcpdocx" -Value $mcpdocxConfig
}

if (-not ($config.PSObject.Properties.Name -contains "inputs")) {
    $config | Add-Member -MemberType NoteProperty -Name "inputs" -Value @()
}

# Hapus mcpServers jika ada (format lama tidak menggunakannya)
if ($config.PSObject.Properties.Name -contains "mcpServers") {
    $config.PSObject.Properties.Remove("mcpServers")
}

# Fungsi untuk format JSON dengan indentasi 2 spasi
function Format-Json {
    param([string]$json)
    
    $indent = 0
    $lines = $json -split '\n'
    $formatted = @()
    
    foreach ($line in $lines) {
        $line = $line.TrimEnd()
        
        # Kurangi indent jika baris dimulai dengan } atau ]
        if ($line -match '^\s*[}\]]') {
            $indent = [Math]::Max(0, $indent - 1)
        }
        
        # Tambah spasi indent
        $formatted += ('  ' * $indent) + $line.TrimStart()
        
        # Tambah indent jika baris diakhiri dengan { atau [
        if ($line -match '[{[]$') {
            $indent++
        }
    }
    
    return $formatted -join "`n"
}

$jsonRaw = $config | ConvertTo-Json -Depth 6
$jsonFormatted = Format-Json -json $jsonRaw
$jsonFormatted | Set-Content -Path $configPath -Encoding UTF8

Write-Info "Konfigurasi MCP disimpan ke $configPath"
Write-Info "Instalasi selesai. Aktivasi venv: .\\.venv\\Scripts\\Activate.ps1"
