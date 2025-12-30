<#
.SYNOPSIS
    Script de remediacion para drivers desactualizados
    
.DESCRIPTION
    Lee el reporte de deteccion de drivers y ejecuta acciones de remediacion.
    - Modo ReportOnly: Solo reporta sin actualizar
    - Modo Update: Intenta actualizar drivers via Windows Update
    Envia resultados a API REST para seguimiento.
    
.PARAMETER Url
    Endpoint de la API para notificacion de resultados (valor predefinido).
.EXAMPLE
    .\Driver_Remediation.ps1
    Ejecuta remediacion de drivers con valores predeterminados.
#>

# URL para notificacion de resultados a API REST
$Url = "http://localhost:8000/api/v1/remediations/"

# Configuracion
$intuneLogsPath = "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs"
$detectionReportPath = Join-Path $intuneLogsPath "DriversStatus.json"
$remediationReportPath = Join-Path $intuneLogsPath "DriversRemediation.json"
$actionMode = "ReportOnly"  # Cambiar a "Update" para intentar actualizaciones

# Verificar si tenemos permisos de escritura, si no usar TEMP
$testPath = Join-Path $intuneLogsPath "test_write_access.tmp"
$useIntunePath = $false
try {
    [System.IO.File]::WriteAllText($testPath, "test", [System.Text.Encoding]::UTF8)
    Remove-Item $testPath -Force -ErrorAction SilentlyContinue
    $useIntunePath = $true
} catch {
    Write-Host "No hay permisos para escribir en $intuneLogsPath" -ForegroundColor Yellow
    Write-Host "Usando directorio temporal alternativo..." -ForegroundColor Yellow
    $detectionReportPath = "$env:TEMP\DriversStatus.json"
    $remediationReportPath = "$env:TEMP\DriversRemediation.json"
}

# Inicializar colecciones de resultados
$Results = [PSCustomObject]@{
    CompletedTasks      = @()   # Tareas completadas exitosamente
    DriversProcessed    = @()   # Drivers procesados
    RemediationDetails  = @()   # Detalles de acciones de remediacion
    CriticalErrors      = @()   # Fallos criticos que detienen procesos
}

function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $adminRole = [Security.Principal.WindowsBuiltInRole]::Administrator
    return ([Security.Principal.WindowsPrincipal]$currentUser).IsInRole($adminRole)
}

function Add-Success {
    param([Parameter(Mandatory=$true)][string]$Task)
    $script:Results.CompletedTasks += $Task
}

function Add-CriticalError {
    param([Parameter(Mandatory=$true)][string]$Message)
    $script:Results.CriticalErrors += $Message
}

function Add-RemediationDetail {
    param(
        [string]$DriverName,
        [string]$Action,
        [string]$Status,
        [string]$Message,
        [string]$OldVersion,
        [string]$NewVersion = ""
    )
    
    $script:Results.RemediationDetails += [PSCustomObject]@{
        Timestamp   = (Get-Date).ToString("HH:mm:ss")
        DriverName  = $DriverName
        Action      = $Action
        Status      = $Status
        Message     = $Message
        OldVersion  = $OldVersion
        NewVersion  = $NewVersion
    }
}

if ($actionMode -eq "Update" -and -not (Test-Administrator)) {
    Write-Host "ERROR: Se requieren privilegios administrativos para actualizar drivers." -ForegroundColor Red
    Add-CriticalError "El script no se ejecuta con privilegios administrativos"
    $actionMode = "ReportOnly"
    Write-Host "Cambiando a modo ReportOnly..." -ForegroundColor Yellow
}

try {
    Write-Host "=== Iniciando proceso de remediacion de drivers ===" -ForegroundColor Cyan
    Write-Host "Modo de operacion: $actionMode" -ForegroundColor White
    Add-Success "Remediacion iniciada: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"

    $computerName = $env:COMPUTERNAME
    $userName = $env:USERNAME
    Add-Success "Equipo: $computerName | Usuario: $userName"

    if (Test-Path $detectionReportPath) {
        $detectionData = Get-Content $detectionReportPath -Raw | ConvertFrom-Json
        $outdatedDrivers = $detectionData.Drivers | Where-Object { $_.NeedsUpdate -eq $true }
        
        Write-Host "Encontrados $($outdatedDrivers.Count) drivers desactualizados en el reporte" -ForegroundColor Cyan
        Add-Success "Reporte de deteccion leido: $($outdatedDrivers.Count) drivers desactualizados"
        
        foreach ($driver in $outdatedDrivers) {
            try {
                Write-Host "`nProcesando: $($driver.DeviceName)" -ForegroundColor Yellow
                Write-Host "  Tipo: $($driver.DriverType)" -ForegroundColor Gray
                Write-Host "  Version actual: $($driver.DriverVersion)" -ForegroundColor Gray
                Write-Host "  Razon: $($driver.UpdateReason)" -ForegroundColor Gray
                
                $Results.DriversProcessed += $driver
                
                if ($actionMode -eq "ReportOnly") {
                    Add-RemediationDetail -DriverName $driver.DeviceName `
                              -Action "Report" `
                              -Status "Reported" `
                              -Message "Driver marcado como desactualizado - $($driver.UpdateReason)" `
                              -OldVersion $driver.DriverVersion
                    
                    Write-Host "  Accion: Reportado (sin actualizar)" -ForegroundColor Blue
                    Add-Success "Driver reportado: $($driver.DeviceName)"
                }
                elseif ($actionMode -eq "Update") {
                    Write-Host "  Intentando actualizacion..." -ForegroundColor Cyan
                    
                    try {
                        $updateSession = New-Object -ComObject Microsoft.Update.Session
                        $updateSearcher = $updateSession.CreateUpdateSearcher()
                        $searchResult = $updateSearcher.Search("Type=''Driver''")
                        
                        if ($searchResult.Updates.Count -gt 0) {
                            $matchingUpdate = $searchResult.Updates | Where-Object {
                                $_.Title -match [regex]::Escape($driver.DeviceName) -or
                                $_.Description -match [regex]::Escape($driver.HardwareID)
                            } | Select-Object -First 1
                            
                            if ($matchingUpdate) {
                                Write-Host "  Actualizacion encontrada: $($matchingUpdate.Title)" -ForegroundColor Green
                                
                                $downloader = $updateSession.CreateUpdateDownloader()
                                $downloader.Updates = $matchingUpdate
                                $downloadResult = $downloader.Download()
                                
                                if ($downloadResult.ResultCode -eq 2) {
                                    $installer = $updateSession.CreateUpdateInstaller()
                                    $installer.Updates = $matchingUpdate
                                    $installationResult = $installer.Install()
                                    
                                    if ($installationResult.ResultCode -eq 2) {
                                        Add-RemediationDetail -DriverName $driver.DeviceName `
                                                  -Action "Update" `
                                                  -Status "Success" `
                                                  -Message "Driver actualizado via Windows Update" `
                                                  -OldVersion $driver.DriverVersion `
                                                  -NewVersion $matchingUpdate.DriverVerification
                                        
                                        Write-Host "  Actualizado correctamente" -ForegroundColor Green
                                        Add-Success "Driver actualizado: $($driver.DeviceName)"
                                    }
                                    else {
                                        throw "Instalacion fallo: $($installationResult.ResultCode)"
                                    }
                                }
                                else {
                                    throw "Descarga fallo: $($downloadResult.ResultCode)"
                                }
                            }
                            else {
                                Write-Host "  No se encontro actualizacion especifica" -ForegroundColor Yellow
                                Add-RemediationDetail -DriverName $driver.DeviceName `
                                          -Action "Update" `
                                          -Status "NoUpdateAvailable" `
                                          -Message "No se encontro actualizacion en Windows Update" `
                                          -OldVersion $driver.DriverVersion
                                Add-Success "Driver sin actualizacion disponible: $($driver.DeviceName)"
                            }
                        }
                        else {
                            Write-Host "  No hay actualizaciones disponibles" -ForegroundColor Yellow
                            Add-RemediationDetail -DriverName $driver.DeviceName `
                                      -Action "Update" `
                                      -Status "NoUpdates" `
                                      -Message "Windows Update no reporto actualizaciones de drivers" `
                                      -OldVersion $driver.DriverVersion
                            Add-Success "No hay actualizaciones para: $($driver.DeviceName)"
                        }
                    }
                    catch {
                        $errorMsg = "Error en Windows Update: $($_.Exception.Message)"
                        Write-Host "  $errorMsg" -ForegroundColor Red
                        Add-RemediationDetail -DriverName $driver.DeviceName `
                                  -Action "Update" `
                                  -Status "Failed" `
                                  -Message $errorMsg `
                                  -OldVersion $driver.DriverVersion
                        Add-CriticalError $errorMsg
                    }
                }
            }
            catch {
                $errorMsg = "Error procesando driver $($driver.DeviceName): $($_.Exception.Message)"
                Write-Host "  $errorMsg" -ForegroundColor Red
                Add-CriticalError $errorMsg
            }
        }
        
        Write-Host "`n=== RESUMEN DE REMEDIACION ===" -ForegroundColor Cyan
        Write-Host "Modo: $actionMode" -ForegroundColor White
        Write-Host "Total drivers procesados: $($Results.DriversProcessed.Count)" -ForegroundColor Gray
        Write-Host "Detalles de remediacion: $($Results.RemediationDetails.Count)" -ForegroundColor Gray
        
        if ($Results.CriticalErrors.Count -gt 0) {
            Write-Host "`nErrores encontrados: $($Results.CriticalErrors.Count)" -ForegroundColor Red
            foreach ($err in $Results.CriticalErrors) {
                Write-Host "  - $err" -ForegroundColor Red
            }
        }
    }
    else {
        Write-Host "No se encontro reporte de deteccion. Ejecute primero el script de deteccion." -ForegroundColor Red
        Add-CriticalError "Reporte de deteccion no encontrado en $detectionReportPath"
    }
}
catch {
    $errorMsg = "Error critico en remediacion: $($_.Exception.Message)"
    Add-CriticalError $errorMsg
    Write-Error $errorMsg
    Write-Host "Tipo de error: $($_.Exception.GetType().Name)" -ForegroundColor Red
}

try {
    Write-Host "`n=== Generando reporte de remediacion ===" -ForegroundColor Cyan
    
    if (Test-Path $remediationReportPath) {
        Remove-Item $remediationReportPath -Force -ErrorAction SilentlyContinue
    }
    
    $localReport = [PSCustomObject]@{
        RemediationTimestamp  = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        ComputerName          = $env:COMPUTERNAME
        UserName              = $env:USERNAME
        ActionMode            = $actionMode
        TotalDriversProcessed = $Results.DriversProcessed.Count
        Results               = $Results
    }
    
    $localReportJson = $localReport | ConvertTo-Json -Depth 10
    
    # Intentar guardar el reporte
    try {
        [System.IO.File]::WriteAllText($remediationReportPath, $localReportJson, [System.Text.Encoding]::UTF8)
        Add-Success "Reporte local guardado en: $remediationReportPath"
        Write-Host "Reporte guardado en: $remediationReportPath" -ForegroundColor Green
    }
    catch {
        # Si falla, intentar en TEMP
        $remediationReportPath = "$env:TEMP\DriversRemediation.json"
        [System.IO.File]::WriteAllText($remediationReportPath, $localReportJson, [System.Text.Encoding]::UTF8)
        Add-Success "Reporte guardado en ubicacion alternativa: $remediationReportPath"
        Write-Host "Reporte guardado en: $remediationReportPath" -ForegroundColor Yellow
    }

    if ($Url -and $Results.DriversProcessed.Count -gt 0) {
        try {
            Write-Host "Enviando resultados a API..." -ForegroundColor Cyan
            
            $deviceId = 1
            $groupId = 1
            
            $bodyObject = @{
                id_group = $groupId
                status = if ($Results.CriticalErrors.Count -eq 0) { $true } else { $false }
                action_remediation = $Results
                id_device = $deviceId
            }
            
            $jsonPayload = $bodyObject | ConvertTo-Json -Depth 10 -Compress
            
            $debugPath = "$env:TEMP\DriverRemediationPayload.json"
            [System.IO.File]::WriteAllText($debugPath, $jsonPayload, [System.Text.Encoding]::UTF8)
            Write-Host "Payload guardado en: $debugPath" -ForegroundColor Gray
            
            $response = Invoke-WebRequest -Uri $Url -Method Post -Body $jsonPayload -ContentType "application/json; charset=utf-8" -ErrorAction Stop
            
            Write-Host "Resultados enviados a API exitosamente (Status: $($response.StatusCode))" -ForegroundColor Green
            Write-Host "Respuesta de API: $($response.Content)" -ForegroundColor White
            Add-Success "Notificacion API enviada exitosamente: $($response.StatusCode)"
            
        } 
        catch {
            Write-Error "Error enviando resultados a API: $($_.Exception.Message)"
            
            if ($_.Exception.Response) {
                try {
                    $errorStream = $_.Exception.Response.GetResponseStream()
                    $reader = New-Object System.IO.StreamReader($errorStream)
                    $errorDetails = $reader.ReadToEnd()
                    Write-Host "Detalles del error HTTP:" -ForegroundColor Red
                    Write-Host $errorDetails -ForegroundColor Red
                    $reader.Dispose()
                    Add-CriticalError "Fallo en notificacion API: $errorDetails"
                }
                catch {
                    Add-CriticalError "Fallo en notificacion API: $($_.Exception.Message)"
                    Write-Host "No se pudieron obtener detalles del error HTTP" -ForegroundColor Yellow
                }
            }
            else {
                Add-CriticalError "Fallo en notificacion API: $($_.Exception.Message)"
            }
        }
    }
}
catch {
    $errorMsg = "Error generando reporte: $($_.Exception.Message)"
    Write-Error $errorMsg
    Add-CriticalError $errorMsg
}

Write-Host "`n=== Finalizacion ===" -ForegroundColor Cyan

if ($Results.CriticalErrors.Count -gt 0) {
    Write-Host "Script completado con errores criticos:" -ForegroundColor Red
    foreach ($crit in $Results.CriticalErrors) {
        Write-Host "  - $crit" -ForegroundColor Red
    }
    exit 1
} 
elseif ($Results.DriversProcessed.Count -gt 0) {
    Write-Host "Script completado: Se procesaron $($Results.DriversProcessed.Count) drivers" -ForegroundColor Yellow
    Write-Host "Modo de operacion: $actionMode" -ForegroundColor Yellow
    exit 0
} 
else {
    Write-Host "Script ejecutado: No hay drivers para procesar" -ForegroundColor Green
    exit 0
}
