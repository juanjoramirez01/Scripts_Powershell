<#
.SYNOPSIS
    Script de detección para problemas de Microsoft Office.
    
.DESCRIPTION
    Este script realiza las siguientes funciones:
    - Verifica el estado del servicio Click-to-Run de Office
    - Comprueba el tamaño de la caché de Microsoft Teams
    - Determina si se requiere ejecutar acciones de remediación

.EXAMPLE
    .\Office_Detection.ps1
    Verifica el estado de los componentes de Office.

.NOTES
    Este script está diseñado para trabajar en conjunto con Office_Remediation.ps1
    como parte de un sistema automatizado de mantenimiento de Office.
#>

<#
.SYNOPSIS
    Verifica si el script se ejecuta con privilegios administrativos.
#>
function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $adminRole = [Security.Principal.WindowsBuiltInRole]::Administrator
    return ([Security.Principal.WindowsPrincipal]$currentUser).IsInRole($adminRole)
}

# Verificar permisos al inicio del script
if (-not (Test-Administrator)) {
    Write-Host "ERROR: Este script requiere privilegios administrativos." -ForegroundColor Red
    Write-Host "Por favor, ejecute el script como Administrador." -ForegroundColor Red
    exit 1
}

try {
    # Inicializar variables de estado
    $needsRemediation = $false
    $remediationReasons = @()
    $TeamsCacheThresholdMB = 100  # Umbral en MB para la caché de Teams

    # 1. Verificar servicio Click-to-Run de Office
    $clickToRunService = Get-Service -Name "ClickToRunSvc" -ErrorAction SilentlyContinue
    
    if ($clickToRunService) {
        if ($clickToRunService.Status -ne 'Running') {
            $needsRemediation = $true
            $remediationReasons += "Servicio Click-to-Run no está ejecutándose (Estado actual: $($clickToRunService.Status))"
        }
    } else {
        Write-Host "Servicio Click-to-Run no encontrado (puede ser normal si Office no está instalado)" -ForegroundColor Yellow
    }

    # 2. Verificar caché de Microsoft Teams
    $teamsCachePaths = @(
        "$env:APPDATA\Microsoft\Teams",
        "$env:LOCALAPPDATA\Microsoft\Teams"
    )
    
    $totalCacheSize = 0
    foreach ($path in $teamsCachePaths) {
        if (Test-Path $path) {
            $cacheSize = (Get-ChildItem $path -Recurse -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum / 1MB
            $totalCacheSize += $cacheSize
        }
    }
    
    if ($totalCacheSize -gt $TeamsCacheThresholdMB) {
        $needsRemediation = $true
        $remediationReasons += "Caché de Teams demasiado grande: $([math]::Round($totalCacheSize, 2)) MB (Umbral: $TeamsCacheThresholdMB MB)"
    }

    # Evaluar si se requiere remediación
    if ($needsRemediation) {
        Write-Host "Se requiere remediación para componentes de Office:" -ForegroundColor Yellow
        foreach ($reason in $remediationReasons) {
            Write-Host "  - $reason" -ForegroundColor Yellow
        }
        exit 1
    } else {
        Write-Host "Todos los componentes de Office están funcionando correctamente. No se requiere remediación." -ForegroundColor Green
        exit 0
    }
}
catch {
    Write-Host "Error al verificar los componentes de Office: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Tipo de error: $($_.Exception.GetType().Name)" -ForegroundColor Red
    if ($_.Exception.InnerException) {
        Write-Host "Error interno: $($_.Exception.InnerException.Message)" -ForegroundColor Red
    }
    exit 1
}
