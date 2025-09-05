<#
.SYNOPSIS
    Script de remediación para problemas comunes de Microsoft Office.
.DESCRIPTION
    Realiza tareas de mantenimiento incluyendo:
    - Reinicio del servicio Click-to-Run de Office
    - Limpieza de la caché de Microsoft Teams
    - Generación de reporte JSON y envío a API REST
    
.PARAMETER Url
    Endpoint de la API para notificación de resultados (valor predefinido).
.EXAMPLE
    .\Office_Remediation.ps1
    Ejecuta todas las tareas de remediación con valores predeterminados.
.NOTES
    Ejecuta tareas de mantenimiento para Microsoft Office y Teams.
#>

# Url para notificación de resultados a API REST
$Url = "http://PSEGBKPGML3.suramericana.com.co:8080/api/remediations"

# Inicializar colecciones de resultados 
$Results = [PSCustomObject]@{
    CompletedTasks   = @()   # Tareas completadas exitosamente
    FileLevelErrors  = @()   # Errores específicos por archivo/directorio
    CriticalErrors   = @()   # Fallos críticos que detienen procesos
}

<#
.SYNOPSIS
    Verifica si el script se ejecuta con privilegios administrativos.
#>
function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $adminRole = [Security.Principal.WindowsBuiltInRole]::Administrator
    return ([Security.Principal.WindowsPrincipal]$currentUser).IsInRole($adminRole)
}

<#
.SYNOPSIS
    Registra una tarea completada exitosamente.
.DESCRIPTION
    Agrega un mensaje de éxito a la colección de resultados.
.PARAMETER Task
    Descripción de la tarea completada.
#>
function Add-Success {
    param([Parameter(Mandatory=$true)][string]$Task)
    $script:Results.CompletedTasks += $Task
}

<#
.SYNOPSIS
    Registra un error a nivel de archivo/directorio.
.DESCRIPTION
    Captura errores durante operaciones con elementos específicos.
.PARAMETER Path
    Ruta del archivo/directorio afectado.
.PARAMETER Message
    Mensaje de error detallado.
#>
function Add-FileError {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][string]$Message
    )
    $script:Results.FileLevelErrors += [PSCustomObject]@{
        Item    = $Path
        Error   = $Message
    }
}

<#
.SYNOPSIS
    Registra un error crítico del sistema.
.DESCRIPTION
    Captura fallos que impiden la ejecución de procesos completos.
.PARAMETER Message
    Descripción del error crítico.
#>
function Add-CriticalError {
    param([Parameter(Mandatory=$true)][string]$Message)
    $script:Results.CriticalErrors += $Message
}

<#
.SYNOPSIS
    Reinicia el servicio Click-to-Run de Office.
.DESCRIPTION
    Detiene y reinicia el servicio Click-to-Run de Office.
#>
function Restart-ClickToRunService {
    try {
        $service = Get-Service -Name "ClickToRunSvc" -ErrorAction Stop
        
        if ($service.Status -eq 'Running') {
            Stop-Service -Name "ClickToRunSvc" -Force -ErrorAction Stop
            Add-Success "Servicio Click-to-Run detenido correctamente"
            
            # Pequeña pausa para asegurar la detención completa
            Start-Sleep -Seconds 3
        }
        
        Start-Service -Name "ClickToRunSvc" -ErrorAction Stop
        Add-Success "Servicio Click-to-Run iniciado correctamente"
        
        # Verificar que el servicio se está ejecutando
        $service = Get-Service -Name "ClickToRunSvc" -ErrorAction Stop
        if ($service.Status -eq 'Running') {
            Add-Success "Servicio Click-to-Run verificado y funcionando correctamente"
        } else {
            Add-CriticalError "El servicio Click-to-Run no se pudo iniciar correctamente"
        }
    }
    catch {
        if ($_.Exception.Message -like "*No se encuentra*") {
            Add-Success "Servicio Click-to-Run no encontrado (puede ser normal si Office no está instalado)"
        } else {
            Add-CriticalError "Error al reiniciar el servicio Click-to-Run: $($_.Exception.Message)"
        }
    }
}

<#
.SYNOPSIS
    Limpia la caché de Microsoft Teams.
.DESCRIPTION
    Elimina archivos temporales y caché de Microsoft Teams.
#>
function Clear-TeamsCache {
    $teamsCachePaths = @(
        "$env:APPDATA\Microsoft\Teams",
        "$env:LOCALAPPDATA\Microsoft\Teams",
        "$env:APPDATA\Microsoft\Teams\Application Cache",
        "$env:APPDATA\Microsoft\Teams\Cache",
        "$env:APPDATA\Microsoft\Teams\blob_storage",
        "$env:APPDATA\Microsoft\Teams\databases",
        "$env:APPDATA\Microsoft\Teams\GPUcache",
        "$env:APPDATA\Microsoft\Teams\IndexedDB",
        "$env:APPDATA\Microsoft\Teams\Local Storage",
        "$env:APPDATA\Microsoft\Teams\tmp"
    )
    
    foreach ($path in $teamsCachePaths) {
        if (Test-Path $path) {
            try {
                # Detener procesos de Teams si están ejecutándose
                $teamsProcesses = Get-Process -Name "Teams" -ErrorAction SilentlyContinue
                if ($teamsProcesses) {
                    Stop-Process -Name "Teams" -Force -ErrorAction SilentlyContinue
                    Start-Sleep -Seconds 3
                }
                
                # Eliminar contenido de la caché
                Remove-Item -Path "$path\*" -Recurse -Force -ErrorAction Stop
                Add-Success "Caché de Teams limpiada: $path"
            }
            catch {
                Add-FileError -Path $path -Message $_.Exception.Message
            }
        } else {
            Add-Success "Directorio de caché de Teams no encontrado: $path"
        }
    }
}

# SECCIÓN PRINCIPAL DE REMEDIACIÓN

# Verificar permisos administrativos
if (-not (Test-Administrator)) {
    Write-Host "ERROR: Este script requiere privilegios administrativos." -ForegroundColor Red
    Write-Host "Por favor, ejecute el script como Administrador." -ForegroundColor Red
    Add-CriticalError "El script no se ejecuta con privilegios administrativos"
    exit 1
}

# 1. Reiniciar el servicio Click-to-Run de Office
Write-Host "Reiniciando servicio Click-to-Run de Office..." -ForegroundColor Yellow
Restart-ClickToRunService

# 2. Limpiar la caché de Microsoft Teams
Write-Host "Limpiando caché de Microsoft Teams..." -ForegroundColor Yellow
Clear-TeamsCache

# GENERACIÓN Y ENVÍO DE RESULTADOS

# Generación de reporte JSON
try {
    if ($env:LOCALAPPDATA) {
        $JsonOutput = $Results | ConvertTo-Json -Depth 4
        [System.IO.File]::WriteAllText("$env:LOCALAPPDATA\OfficeRemediationResults.json", $JsonOutput, [System.Text.Encoding]::UTF8)
        
        Write-Output $JsonOutput
    } else {
        Add-CriticalError "Variable LOCALAPPDATA no encontrada"
        $JsonOutput = $Results | ConvertTo-Json -Depth 4
        Write-Output $JsonOutput
        try {
            [System.IO.File]::WriteAllText("OfficeRemediationResults.json", $JsonOutput, [System.Text.Encoding]::UTF8)
        } catch {
            Write-Warning "Error guardando resultados: $($_.Exception.Message)"
            Add-CriticalError "Error guardando resultados: $($_.Exception.Message)"
        }
    }
}
catch {
    Write-Error "Fallo generando salida: $($_.Exception.Message)"
    Add-CriticalError "Fallo generando salida: $($_.Exception.Message)"
}

# Notificación a API REST (separada de la generación del JSON)
try {
    $deviceId = 1
    $groupId = "Auto-created Group: OFFICE_REMEDIATION_$(Get-Date -Format 'yyyyMMdd')"
    
    $body = @{
        id_group = $groupId
        status = 1
        action_remediation = $Results
        id_device = $deviceId
    } | ConvertTo-Json -Depth 5
    
    $response = Invoke-WebRequest -Uri $Url -Method POST -Body $body -Headers @{"Accept"="application/json"; "Content-Type"="application/json"} -ErrorAction Stop
    
    Add-Success "Notificación API enviada: $($response.StatusCode)"
    
} catch {
    $errorMessage = "Fallo en notificación API"
    
    # Capturar detalles del error HTTP
    if ($_.Exception.Response) {
        try {
            $errorStream = $_.Exception.Response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($errorStream)
            try {
                $errorDetails = $reader.ReadToEnd()
                $errorMessage += ": $errorDetails"
            }
            finally {
                $reader.Dispose()
            }
        } catch {
            $errorMessage += ": $($_.Exception.Message)"
        }
    } else {
        $errorMessage += ": $($_.Exception.Message)"
    }
    
    # Agregar como error de archivo en lugar de error crítico para no afectar el resultado principal
    Add-FileError -Path "API Notification" -Message $errorMessage
    Write-Warning "Error en notificación API: $errorMessage"
}

# MANEJO FINAL DE ESTADO
# Solo verificar errores críticos de las operaciones principales, no de la notificación API
if ($Results.CriticalErrors.Count -gt 0) {
    Write-Host "`nScript completado con errores críticos. Verifique registros." -ForegroundColor Red
    exit 1
} else {
    Write-Host "`nScript ejecutado exitosamente." -ForegroundColor Green
    exit 0
}
