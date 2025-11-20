# ======================
# 1. CPU
# ======================
Write-Host "`n=== CPU ==="

Get-CimInstance Win32_Processor |
    Select-Object Name, NumberOfCores, NumberOfLogicalProcessors | Format-List

# ======================
# 2. RAM usando SMBIOSMemoryType (más preciso)
# ======================
Write-Host "`n=== RAM ==="


Get-CimInstance Win32_PhysicalMemory | ForEach-Object {
    $type = switch ($_.SMBIOSMemoryType) {
        20 {"DDR"}
        21 {"DDR2"}
        24 {"DDR3"}
        26 {"DDR4"}
        34 {"DDR5"}
        default {"Desconocido"}
    }

    [PSCustomObject]@{
        Capacidad_GB = "{0:N2}" -f ($_.Capacity / 1GB)
        Tipo_DDR = $type
        Frecuencia_MHz = $_.ConfiguredClockSpeed
    }
} | Format-Table -AutoSize

# ======================
# 3. Activación Windows
# ======================
Write-Host "`nAbriendo ventana de activación de Windows..."
Start-Process "slmgr.vbs" -ArgumentList "/xpr"

# ======================
# 4. Office activado (limpio)
# ======================
Write-Host "`n=== OFFICE ==="

$officePaths = @(
    "C:\Program Files\Microsoft Office\Office16\OSPP.VBS",
    "C:\Program Files (x86)\Microsoft Office\Office16\OSPP.VBS",
    "C:\Program Files\Microsoft Office\Office15\OSPP.VBS",
    "C:\Program Files (x86)\Microsoft Office\Office15\OSPP.VBS"
)

$ospp = $officePaths | Where-Object { Test-Path $_ } | Select-Object -First 1

if ($ospp) {
    Write-Host "Office detectado. Verificando activación..."
    $result = cscript.exe //nologo $ospp /dstatus
    $result | Select-String "LICENSE STATUS"
} else {
    Write-Host "No se encontró OSPP.VBS. Abriendo Word..."
    Start-Process "winword.exe"
}

# ======================
# 5. Administración de discos
# ======================
Write-Host "`nAbriendo administración de discos..."
Start-Process "diskmgmt.msc"


# ======================
# 6. Drivers faltantes
# ======================
Write-Host "`n=== Drivers con error ==="

try {
    $bad = Get-PnpDevice -Status Error -ErrorAction SilentlyContinue
    if ($bad) {
        $bad | Select-Object Name, Status, ProblemCode
    } else {
        Write-Host "No hay dispositivos con errores." -ForegroundColor Green
    }
} catch {
    Write-Host "No disponible en esta versión de Windows."
}


# ======================
# 7. Detectar Windows 10 vs Windows 11 (método correcto)
# ======================
$os = Get-CimInstance Win32_OperatingSystem

$build = [int]$os.BuildNumber
if ($build -ge 22000) {
    $winver = "Windows 11"
} else {
    $winver = "Windows 10"
}

Write-Host "`n=== Sistema ==="
Write-Host "Sistema detectado: $winver"
Write-Host "Build: $build"
Write-Host "Edición: $($os.Caption)"
Write-Host "Arquitectura: $($os.OSArchitecture)"

Write-Host "`nAbriendo Propiedades del sistema..."
Start-Process "ms-settings:about"

# ======================
# 8. CrystalDiskInfo
# ======================
Write-Host ""
Write-Host "Abriendo CrystalDiskInfo en modo privado..."


$edge = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"

$url = "https://crystalmark.info/en/software/crystaldiskinfo/"

if (Test-Path $edge) {
    Start-Process $edge -ArgumentList "--inprivate $url"
} else {
    Start-Process $url
}

# ======================
# 9. Páginas de testeo
# ======================
Write-Host ""
Write-Host "Abriendo paginas de testeo..."


$testSites = @(
    "https://es.mictests.com/",
    "https://www.onlinemictest.com/es/prueba-de-teclado/",
    "https://es.webcamtests.com/"
)

foreach ($site in $testSites) {
    if (Test-Path $edge) {
        Start-Process $edge -ArgumentList "--inprivate $site"
    } else {
        Start-Process $site
    }
}

Write-Host "`n=== FIN DEL INFORME ===`n"
