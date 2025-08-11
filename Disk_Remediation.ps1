<#
.SYNOPSIS
    Script de remediación para liberar espacio en disco mediante limpieza automatizada.
.DESCRIPTION
    Realiza tareas de mantenimiento incluyendo:
    - Eliminación de archivos temporales del sistema y usuario
    - Limpieza de caché de navegadores (Chrome, Edge)
    - Eliminación de perfiles de usuario inactivos (>6 meses)
    - Generación de reporte JSON y envío a API REST
    
.PARAMETER Url
    Endpoint de la API para notificación de resultados (valor predefinido).
.EXAMPLE
    .\Disk_Remediation.ps1
    Ejecuta todas las tareas de limpieza con valores predeterminados.
#>

# Url para notificación de resultados a API REST
$Url = "http://PSEGBKPGML3.suramericana.com.co:8080/api/remediations"

# Configuración de inactividad (perfiles no usados en los últimos 6 meses)
[int]$monthsInactive = 6
$cutoffDate = (Get-Date).AddMonths(-$monthsInactive)

# Inicializar colecciones de resultados 
$Results = [PSCustomObject]@{
    CompletedTasks   = @()   # Tareas completadas exitosamente
    FileLevelErrors  = @()   # Errores específicos por archivo/directorio
    CriticalErrors   = @()   # Fallos críticos que detienen procesos
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
    Elimina contenido de directorios de forma segura.
.DESCRIPTION
    Borra recursivamente el contenido de rutas especificadas con manejo de errores.
.PARAMETER PathPattern
    Patrón de ruta (ej: C:\Temp\*).
#>
function Clear-Directory {
    param([Parameter(Mandatory=$true)][string]$PathPattern)

    try {
        $basePath = Split-Path $PathPattern -Parent
        # Validar existencia de ruta base
        if (-not (Test-Path $basePath -ErrorAction SilentlyContinue)) {
            Add-FileError -Path $PathPattern -Message "La ruta base no existe o no es accesible"
            return
        }
        
        # Recuperar elementos con manejo de errores
        $items = Get-ChildItem -Path $PathPattern -Recurse -Force -ErrorAction Stop
        
        # Procesar cada elemento individualmente
        foreach ($item in $items) {
            try {
                Remove-Item -LiteralPath $item.FullName -Recurse -Force -ErrorAction Stop
                Add-Success "Eliminado: $($item.FullName)"
            }
            catch {
                Add-FileError -Path $item.FullName -Message $_.Exception.Message
            }
        }
    }
    catch {
        Add-CriticalError "Fallo al enumerar la ruta '$PathPattern': $($_.Exception.Message)"
    }
}

# SECCIÓN PRINCIPAL DE LIMPIEZA

# Limpieza de temporales del sistema
try {
    if ($env:SystemRoot) {
        Clear-Directory -Path "$env:SystemRoot\Temp\*"
    } else {
        Add-CriticalError "Variable de entorno SystemRoot no encontrada"
    }
}
catch {
    Add-CriticalError "Error en limpieza de temporales del sistema: $($_.Exception.Message)"
}

# Limpieza de temporales del usuario
try {
    if ($env:TEMP) {
        Clear-Directory -Path "$env:TEMP\*"
    } elseif (Test-Path "c:\temp") {
        Clear-Directory -Path "c:\temp\*"
    } else {
        Add-CriticalError "Variable TEMP no existe y 'c:\temp' no encontrado"
    }
}
catch {
    Add-CriticalError "Error en limpieza de temporales de usuario: $($_.Exception.Message)"
}

# Limpieza de caché de navegadores
try {
    if ($env:USERPROFILE) {
        $browserPaths = @(
            "$env:USERPROFILE\AppData\Local\Google\Chrome\User Data\Default\Cache\*",
            "$env:USERPROFILE\AppData\Local\Microsoft\Edge\User Data\Default\Cache\*"
        )

        foreach ($path in $browserPaths) {
            try {
                Clear-Directory -Path $path
            }
            catch {
                Add-CriticalError "Error limpiando caché en '$path': $($_.Exception.Message)"
            }
        }
    } else {
        Add-CriticalError "Variable USERPROFILE no encontrada"
    }
}
catch {
    Add-CriticalError "Error inicializando limpieza de navegadores: $($_.Exception.Message)"
}

# Eliminación de perfiles inactivos
try {
    # Consultar perfiles no utilizados en últimos 6 meses
    $oldProfiles = Get-CimInstance -Class Win32_UserProfile -ErrorAction Stop | Where-Object {
        -not $_.Special -and 
        $_.Loaded -eq $false -and
        $_.LastUseTime -and
        $_.LastUseTime -ne $null -and
        ([Management.ManagementDateTimeConverter]::ToDateTime($_.LastUseTime) -lt $cutoffDate)
    }

    Add-Success "Perfiles inactivos detectados: $($oldProfiles.Count)"

    # Procesar cada perfil
    foreach ($prof in $oldProfiles) {
        $path = $prof.LocalPath
        
        # Validar rutas seguras para eliminación
        if ($path -and (Test-Path -Path $path -ErrorAction SilentlyContinue) -and $path -notmatch '^C:\\(Windows|Program Files)') {
            try {
                $prof | Remove-CimInstance -ErrorAction Stop
                
                # Eliminación física de directorio
                if (Test-Path -Path $path -ErrorAction SilentlyContinue) {
                    Remove-Item -Path $path -Recurse -Force -ErrorAction Stop
                    Add-Success "Perfil eliminado: $path"
                }
            }
            catch {
                Add-FileError -Path $path -Message $_.Exception.Message
            }
        } else {
            # Manejo de casos especiales
            if (-not $path) {
                Add-FileError -Path "Desconocido" -Message "Ruta de perfil vacía"
            } elseif (-not (Test-Path -Path $path -ErrorAction SilentlyContinue)) {
                Add-FileError -Path $path -Message "Ruta no existe"
            } else {
                Add-FileError -Path $path -Message "Ruta protegida (directorio del sistema)"
            }
        }
    }
}
catch {
    Add-CriticalError "Error detectando/eliminando perfiles: $($_.Exception.Message)"
}

# GENERACIÓN Y ENVÍO DE RESULTADOS

# Generación de reporte JSON
try {
    if ($env:LOCALAPPDATA) {
        $JsonOutput = $Results | ConvertTo-Json -Depth 4
        [System.IO.File]::WriteAllText("$env:LOCALAPPDATA\CleanupResults.json", $JsonOutput, [System.Text.Encoding]::UTF8)
        
        Write-Output $JsonOutput
        
        # Notificación a API REST
        try {
            $deviceId = 1
            $groupId = "Auto-created Group: DISK_CLEANUP_20250731"
            
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
            
            Add-CriticalError $errorMessage
        }
    } else {
        Add-CriticalError "Variable LOCALAPPDATA no encontrada"
        $JsonOutput = $Results | ConvertTo-Json -Depth 4
        Write-Output $JsonOutput
        try {
            [System.IO.File]::WriteAllText("CleanupResults.json", $JsonOutput, [System.Text.Encoding]::UTF8)
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

# MANEJO FINAL DE ESTADO
if ($Results.CriticalErrors.Count -gt 0) {
    Write-Host "Script completado con errores críticos. Verifique registros." -ForegroundColor Red
    exit 1
} else {
    Write-Host "Script ejecutado exitosamente." -ForegroundColor Green
    exit 0
}
