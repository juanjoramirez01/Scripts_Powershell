<#
.SYNOPSIS
    Script de remediación para solucionar problemas del servicio Print Spooler.
.DESCRIPTION
    Realiza tareas de mantenimiento incluyendo:
    - Reinicio del servicio Print Spooler
    - Eliminación de trabajos de impresión atascados
    - Limpieza de archivos temporales de impresión
    - Generación de reporte JSON y envío a API REST
    
.PARAMETER Url
    Endpoint de la API para notificación de resultados (valor predefinido).
.EXAMPLE
    .\Spooler_Remediation.ps1
    Ejecuta todas las tareas de remediación con valores predeterminados.
#>

# Url para notificación de resultados a API REST
$Url = "http://PSEGBKPGML3.suramericana.com.co:8080/api/remediations"

# Inicializar colecciones de resultados 
$Results = [PSCustomObject]@{
    CompletedTasks   = @()   # Tareas completadas exitosamente
    FileLevelErrors  = @()   # Errores específicos por archivo/trabajo
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
    Registra un error a nivel de archivo/trabajo.
.DESCRIPTION
    Captura errores durante operaciones con elementos específicos.
.PARAMETER Path
    Ruta del archivo/trabajo afectado.
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
    Verifica y repara problemas con WMI si es necesario.
#>
function Repair-WMI {
    try {
        # Verificar el estado del servicio WMI
        $winmgmtService = Get-Service -Name Winmgmt -ErrorAction Stop
        
        if ($winmgmtService.Status -ne 'Running') {
            Write-Host "Iniciando servicio WMI..." -ForegroundColor Yellow
            Start-Service -Name Winmgmt -ErrorAction Stop
            Add-Success "Servicio WMI iniciado correctamente"
        }
        
        # Pequeña pausa para asegurar la inicialización completa
        Start-Sleep -Seconds 3
        
        # Verificar namespaces de impresión
        $printNamespace = Get-CimInstance -Namespace root/cimv2 -ClassName __Namespace -Filter "Name='Printing'" -ErrorAction SilentlyContinue
        
        if (-not $printNamespace) {
            Write-Host "Advertencia: Namespace de impresión no encontrado en WMI" -ForegroundColor Yellow
            return $false
        }
        
        Add-Success "WMI verificado correctamente"
        return $true
    }
    catch {
        $errorMsg = "Error al verificar/reparar WMI: $($_.Exception.Message)"
        if ($_.Exception.InnerException) {
            $errorMsg += " | Error interno: $($_.Exception.InnerException.Message)"
        }
        
        Add-CriticalError $errorMsg
        return $false
    }
}

<#
.SYNOPSIS
    Elimina trabajos de impresión atascados.
.DESCRIPTION
    Elimina trabajos de impresión con más de 1 hora de antigüedad.
#>
function Clear-StuckPrintJobs {
    try {
        # Obtener trabajos de impresión atascados usando CIM
        $stuckJobs = Get-CimInstance -ClassName Win32_PrintJob -ErrorAction Stop | Where-Object {
            $_.TimeSubmitted -and 
            ([Management.ManagementDateTimeConverter]::ToDateTime($_.TimeSubmitted) -lt (Get-Date).AddHours(-1))
        }

        if ($stuckJobs -and $stuckJobs.Count -gt 0) {
            Add-Success "Trabajos atascados detectados: $($stuckJobs.Count)"
            
            # Eliminar cada trabajo atascado
            foreach ($job in $stuckJobs) {
                try {
                    # Intentar eliminar usando CIM
                    Remove-CimInstance -CimInstance $job -ErrorAction Stop
                    Add-Success "Trabajo eliminado: $($job.JobId) en $($job.Name)"
                }
                catch {
                    $errorDetails = $_.Exception.Message
                    # Intentar método alternativo si CIM falla
                    try {
                        Write-Host "Intentando método alternativo para eliminar trabajo $($job.JobId)..." -ForegroundColor Yellow
                        Invoke-CimMethod -InputObject $job -MethodName Delete -ErrorAction Stop
                        Add-Success "Trabajo eliminado (método alternativo): $($job.JobId)"
                    }
                    catch {
                        $errorMsg = "Error eliminando trabajo $($job.JobId): $errorDetails"
                        if ($_.Exception.Message) {
                            $errorMsg += " | Método alternativo también falló: $($_.Exception.Message)"
                        }
                        Add-FileError -Path "Trabajo $($job.JobId)" -Message $errorMsg
                    }
                }
            }
        } else {
            Add-Success "No se encontraron trabajos de impresión atascados"
        }
    }
    catch {
        # Capturar información detallada del error
        $errorMsg = "Error al obtener/eliminar trabajos de impresión: $($_.Exception.Message)"
        if ($_.Exception.InnerException) {
            $errorMsg += " | Error interno: $($_.Exception.InnerException.Message)"
        }
        
        # Información adicional para diagnóstico
        $errorMsg += " | Categoría: $($_.CategoryInfo.Category) | Estado: $($_.CategoryInfo.Reason)"
        
        Add-CriticalError $errorMsg
    }
}

<#
.SYNOPSIS
    Limpia archivos temporales de impresión.
.DESCRIPTION
    Elimina archivos temporales del directorio spool de impresión.
#>
function Clear-SpoolFiles {
    try {
        $spoolPath = "$env:SystemRoot\System32\spool\PRINTERS"
        
        if (Test-Path $spoolPath) {
            $files = Get-ChildItem -Path $spoolPath -Filter *.spl -ErrorAction Stop
            
            if ($files -and $files.Count -gt 0) {
                Add-Success "Archivos temporales detectados: $($files.Count)"
                
                foreach ($file in $files) {
                    try {
                        Remove-Item -LiteralPath $file.FullName -Force -ErrorAction Stop
                        Add-Success "Archivo eliminado: $($file.Name)"
                    }
                    catch {
                        Add-FileError -Path $file.FullName -Message $_.Exception.Message
                    }
                }
            } else {
                Add-Success "No se encontraron archivos temporales en el directorio spool"
            }
        } else {
            Add-FileError -Path $spoolPath -Message "Directorio spool no encontrado"
        }
    }
    catch {
        Add-CriticalError "Error al limpiar archivos temporales: $($_.Exception.Message)"
    }
}

<#
.SYNOPSIS
    Reinicia el servicio Print Spooler.
.DESCRIPTION
    Detiene y reinicia el servicio Print Spooler de manera controlada.
#>
function Restart-SpoolerService {
    try {
        # Detener el servicio
        Stop-Service -Name Spooler -Force -ErrorAction Stop
        Add-Success "Servicio Print Spooler detenido correctamente"
        
        # Pequeña pausa para asegurar la detención completa
        Start-Sleep -Seconds 5
        
        # Iniciar el servicio
        Start-Service -Name Spooler -ErrorAction Stop
        Add-Success "Servicio Print Spooler iniciado correctamente"
        
        # Pequeña pausa para asegurar la inicialización completa
        Start-Sleep -Seconds 3
        
        # Verificar que el servicio se está ejecutando
        $service = Get-Service -Name Spooler -ErrorAction Stop
        if ($service.Status -eq 'Running') {
            Add-Success "Servicio Print Spooler verificado y funcionando correctamente"
        } else {
            Add-CriticalError "El servicio Print Spooler no se pudo iniciar correctamente"
        }
    }
    catch {
        Add-CriticalError "Error al reiniciar el servicio Print Spooler: $($_.Exception.Message)"
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

# Verificar y reparar WMI primero
Write-Host "Verificando estado de WMI..." -ForegroundColor Yellow
$wmiHealthy = Repair-WMI

if ($wmiHealthy) {
    # Limpiar trabajos de impresión atascados
    Write-Host "Buscando trabajos de impresión atascados..." -ForegroundColor Yellow
    Clear-StuckPrintJobs
} else {
    Add-CriticalError "No se pueden limpiar trabajos de impresión debido a problemas con WMI"
}

# Limpiar archivos temporales de impresión
Write-Host "Limpiando archivos temporales de impresión..." -ForegroundColor Yellow
Clear-SpoolFiles

# Reiniciar el servicio Print Spooler
Write-Host "Reiniciando el servicio Print Spooler..." -ForegroundColor Yellow
Restart-SpoolerService

# GENERACIÓN Y ENVÍO DE RESULTADOS

# Generación de reporte JSON
try {
    if ($env:LOCALAPPDATA) {
        $JsonOutput = $Results | ConvertTo-Json -Depth 4
        [System.IO.File]::WriteAllText("$env:LOCALAPPDATA\SpoolerCleanupResults.json", $JsonOutput, [System.Text.Encoding]::UTF8)
        
        Write-Output $JsonOutput
    } else {
        Add-CriticalError "Variable LOCALAPPDATA no encontrada"
        $JsonOutput = $Results | ConvertTo-Json -Depth 4
        Write-Output $JsonOutput
        try {
            [System.IO.File]::WriteAllText("SpoolerCleanupResults.json", $JsonOutput, [System.Text.Encoding]::UTF8)
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
    $groupId = "Auto-created Group: SPOOLER_CLEANUP_$(Get-Date -Format 'yyyyMMdd')"
    
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
    Write-Host "Script completado con errores críticos. Verifique registros." -ForegroundColor Red
    exit 1
} else {
    Write-Host "Script ejecutado exitosamente." -ForegroundColor Green
    exit 0
}
