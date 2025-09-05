<#
.SYNOPSIS
    Script de detección para monitorear el estado del servicio de cola de impresión.
    
.DESCRIPTION
    Este script realiza las siguientes funciones:
    - Verifica el estado del servicio de cola de impresión
    - Comprueba si hay trabajos de impresión atascados
    - Determina si se requiere ejecutar acciones de remediación

.EXAMPLE
    .\Spooler_Detection.ps1
    Verifica el estado del servicio de cola de impresión y trabajos atascados.

.NOTES
    Este script está diseñado para trabajar en conjunto con Spooler_Remediation.ps1
    como parte de un sistema automatizado de mantenimiento del servicio de impresión.
#>

try {
    # Obtener información del servicio de cola de impresión
    $spoolerService = Get-Service -Name Spooler -ErrorAction Stop

    # Validar que se obtuvo información del servicio
    if (-not $spoolerService) {
        Write-Host "Error: No se pudo obtener información del servicio de cola de impresión" -ForegroundColor Red
        exit 1
    }

    # Inicializar variables de estado
    $needsRemediation = $false
    $remediationReasons = @()

    # Verificar estado del servicio
    if ($spoolerService.Status -ne 'Running') {
        $needsRemediation = $true
        $remediationReasons += "Servicio no está ejecutándose (Estado actual: $($spoolerService.Status))"
    }

    # Verificar trabajos atascados si está configurado
    if ($spoolerService.Status -eq 'Running') {
        try {
            # Obtener trabajos de impresión con más de 1 hora de antigüedad usando CIM
            $stuckJobs = Get-CimInstance -ClassName Win32_PrintJob -ErrorAction Stop | Where-Object {
                $_.TimeSubmitted -and 
                ([Management.ManagementDateTimeConverter]::ToDateTime($_.TimeSubmitted) -lt (Get-Date).AddHours(-1))
            }

            if ($stuckJobs -and $stuckJobs.Count -gt 0) {
                $needsRemediation = $true
                $remediationReasons += "Se encontraron $($stuckJobs.Count) trabajos de impresión atascados"
            }
        }
        catch {
            Write-Host "Advertencia: No se pudieron verificar los trabajos de impresión: $($_.Exception.Message)" -ForegroundColor Yellow
            # No salimos con error ya que el servicio podría estar funcionando correctamente
        }
    }

    # Evaluar si se requiere remediación
    if ($needsRemediation) {
        Write-Host "El servicio de cola de impresión requiere remediación:" -ForegroundColor Yellow
        foreach ($reason in $remediationReasons) {
            Write-Host "  - $reason" -ForegroundColor Yellow
        }
        exit 1
    } else {
        Write-Host "El servicio de cola de impresión está funcionando correctamente. No se requiere remediación." -ForegroundColor Green
        Write-Host "Estado del servicio: $($spoolerService.Status)" -ForegroundColor Green
        exit 0
    }
}
catch {
    Write-Host "Error al verificar el servicio de cola de impresión: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Tipo de error: $($_.Exception.GetType().Name)" -ForegroundColor Red
    if ($_.Exception.InnerException) {
        Write-Host "Error interno: $($_.Exception.InnerException.Message)" -ForegroundColor Red
    }
    exit 1
}
