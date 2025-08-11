<#
.SYNOPSIS
    Script de detección para monitorear el uso del espacio en disco y determinar la necesidad de remediación.
    
.DESCRIPTION
    Este script realiza las siguientes funciones:
    - Verifica el porcentaje de uso del disco C:
    - Compara el uso actual contra un umbral configurable
    - Determina si se requiere ejecutar acciones de remediación
    - Proporciona códigos de salida para automatización
    
.PARAMETER Threshold
    Porcentaje de uso del disco que activa la necesidad de remediación.
    Valor predeterminado: 90%

.EXAMPLE
    .\Disk_Detection.ps1
    Verifica el uso del disco C: usando el umbral predeterminado del 90%.

.NOTES
    Este script está diseñado para trabajar en conjunto con Disk_Remediation.ps1
    como parte de un sistema automatizado de mantenimiento de disco.
#>

# Umbral de espacio usado para iniciar la remediación
# Este script verifica el uso del espacio en disco y realiza acciones de remediación si es necesario

param(
    [int]$Threshold = 90
)

try {
    # Obtener información del disco C: usando CIM para mayor compatibilidad
    $disks = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'" -ErrorAction Stop

    # Validar que se obtuvo información del disco
    if (-not $disks) {
        Write-Host "Error: No se pudo obtener información del disco C:" -ForegroundColor Red
        exit 1
    }

    # Calcular el porcentaje de espacio utilizado
    $usedPercent = (($disks.Size - $disks.FreeSpace) / $disks.Size) * 100

    # Evaluar si se requiere remediación basado en el umbral
    if ($usedPercent -ge $Threshold) {
        Write-Host "El espacio en disco C: está al $([math]::Round($usedPercent, 2))% usado. Iniciando remediación..." -ForegroundColor Yellow
        Write-Host "Umbral configurado: $Threshold%" -ForegroundColor Yellow
        Write-Host "Espacio libre: $([math]::Round($disks.FreeSpace / 1GB, 2)) GB de $([math]::Round($disks.Size / 1GB, 2)) GB total" -ForegroundColor Yellow
        exit 1
    } else {
        Write-Host "El espacio en disco C: está al $([math]::Round($usedPercent, 2))% usado. No se requiere remediación." -ForegroundColor Green
        Write-Host "Umbral configurado: $Threshold%" -ForegroundColor Green
        Write-Host "Espacio libre: $([math]::Round($disks.FreeSpace / 1GB, 2)) GB de $([math]::Round($disks.Size / 1GB, 2)) GB total" -ForegroundColor Green
        exit 0
    }
}
catch {
    Write-Host "Error al verificar el espacio en disco: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Tipo de error: $($_.Exception.GetType().Name)" -ForegroundColor Red
    exit 1
}
