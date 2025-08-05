# Umbral de espacio usado para iniciar la remediación
# Este script verifica el uso del espacio en disco y realiza acciones de remediación si es necesario

param(
    [int]$Threshold = 20
)

try {
    $disks = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'"

    $usedPercent = (($disks.Size - $disks.FreeSpace) / $disks.Size) * 100

    if ($usedPercent -ge $Threshold) {
        Write-Host "El espacio en disco C: está al $([math]::Round($usedPercent, 2))% usado. Iniciando remediación..." -ForegroundColor Yellow
        exit 1
    } else {
        Write-Host "El espacio en disco C: está al $([math]::Round($usedPercent, 2))% usado. No se requiere remediación." -ForegroundColor Green
        exit 0
    }
}
catch {
    Write-Host "Error al verificar el espacio en disco: $_" -ForegroundColor Red
    exit 1
}