<#
.SYNOPSIS
    Script de remediacion para actualizacion de drivers criticos del sistema.
.DESCRIPTION
    Realiza tareas de mantenimiento de drivers incluyendo:
    - Deteccion de drivers desactualizados (video, audio, red, impresion)
    - Analisis de versiones y antiguedad
    - Identificacion de drivers no firmados
    - Generacion de reporte JSON y envio a API REST
    
.PARAMETER Url
    Endpoint de la API para notificacion de resultados (valor predefinido).
.EXAMPLE
    .\Driver_Remediation.ps1
    Ejecuta analisis de drivers con valores predeterminados.
#>

# Url para notificacion de resultados a API REST
$Url = "http://localhost:8000/api/v1/remediations/"

# Configuracion de analisis de drivers
$outputPath = "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs\DriversStatus.json"

# Tipos de drivers a monitorear
$targetDeviceClasses = @(
    'Display',       # Video
    'MEDIA',         # Audio
    'Sound',         # Audio alternativo
    'Net',           # Red
    'NetClient',     # Cliente de red
    'NetService',    # Servicios de red
    'NetTrans',      # Transporte de red
    'PrintQueue'     # Cola de impresión
)

# Inicializar colecciones de resultados 
$Results = [PSCustomObject]@{
    CompletedTasks   = @()   # Tareas completadas exitosamente
    DriversAnalyzed  = @()   # Informacion detallada de drivers
    OutdatedDrivers  = @()   # Drivers que requieren actualizacion
    CriticalErrors   = @()   # Fallos criticos que detienen procesos
}

# Funcion para validar formato DMTF antes de convertirlo
function Convert-DmtfSafe {
    param(
        [string]$DmtfString
    )

    # Formato DMTF válido: 14 dígitos . 6 dígitos + 3 dígitos de zona
    if ($DmtfString -and $DmtfString -match '^\d{14}\.\d{6}[\+\-]\d{3}$') {
        try {
            return [Management.ManagementDateTimeConverter]::ToDateTime($DmtfString)
        }
        catch {
            return $null
        }
    }

    return $null
}

try {
    # Obtener todos los drivers del sistema
    Write-Host "Obteniendo informacion de drivers del sistema..." -ForegroundColor Cyan

    # Metodo 1: Usando Win32_PnPSignedDriver (mas rapido)
    $allDrivers = Get-CimInstance -ClassName Win32_PnPSignedDriver -ErrorAction Stop | Where-Object {
        $_.DeviceClass -in $targetDeviceClasses -and
        $_.DriverVersion -ne $null -and
        $_.DriverVersion -ne ""
    }

    # Si no hay resultados con CIM, intentar con WMI
    if (-not $allDrivers -or $allDrivers.Count -eq 0) {
        Write-Host "Intentando con WMI como alternativa..." -ForegroundColor Yellow
        $allDrivers = Get-WmiObject -Class Win32_PnPSignedDriver -ErrorAction Stop | Where-Object {
            $_.DeviceClass -in $targetDeviceClasses -and
            $_.DriverVersion -ne $null -and
            $_.DriverVersion -ne ""
        }
    }

    # Analizar cada driver critico
    foreach ($driver in $allDrivers) {
        try {
            # Determinar tipo de driver
            $driverType = switch ($driver.DeviceClass) {
                'Display' { 'Video' }
                { $_ -in 'MEDIA', 'Sound' } { 'Audio' }
                { $_ -in 'Net', 'NetClient', 'NetService', 'NetTrans' } { 'Red' }
                'PrintQueue' { 'Impresion' }
                default { 'Otro' }
            }

            # Calcular antiguedad del driver
            $driverDateObj = Convert-DmtfSafe $driver.DriverDate
            if ($driverDateObj) {
                $driverAge = (New-TimeSpan -Start $driverDateObj -End (Get-Date)).Days
            } else {
                $driverAge = "Desconocido"
            }

            # Crear objeto de informacion del driver
            $driverInfo = [PSCustomObject]@{
                DeviceName     = $driver.DeviceName
                DeviceClass    = $driver.DeviceClass
                DriverType     = $driverType
                DriverVersion  = $driver.DriverVersion
                DriverDate     = if ($driverDateObj) { $driverDateObj.ToString("yyyy-MM-dd") } else { "Desconocido" }
                DriverAgeDays  = $driverAge
                Manufacturer   = $driver.Manufacturer
                HardwareID     = $driver.HardwareID
                IsSigned       = if ($driver.IsSigned -eq $null) { "Desconocido" } else { [bool]$driver.IsSigned }
                NeedsUpdate    = $false
                UpdateReason   = ""
                Timestamp      = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            }

            # Logica de deteccion de desactualizacion (personalizable)
            $needsUpdate = $false
            $updateReason = @()

            # Regla 1: Driver muy antiguo (> 2 años)
            if ($driverAge -ne $null -and $driverAge -is [int] -and $driverAge -gt 730) {
                $needsUpdate = $true
                $updateReason += "Driver muy antiguo ($driverAge dias)"
            }

            # Regla 2: Version especifica conocida problematica (ejemplo para NVIDIA)
            if ($driver.DeviceName -match "NVIDIA" -and $driver.DriverVersion -match "^2[0-9]") {
                $needsUpdate = $true
                $updateReason += "Version NVIDIA conocida con problemas"
            }

            # Regla 3: Driver no firmado (solo alerta)
            if ($driverInfo.IsSigned -eq $false) {
                $updateReason += "Driver no firmado (potencial riesgo)"
            }

            # Regla 4: Driver de impresion con mas de 1 año
            if ($driverType -eq 'Impresion' -and $driverAge -gt 365) {
                $needsUpdate = $true
                $updateReason += "Driver de impresion desactualizado"
            }

            # Actualizar estado
            $driverInfo.NeedsUpdate = $needsUpdate
            $driverInfo.UpdateReason = if ($updateReason.Count -gt 0) { $updateReason -join "; " } else { "" }

            # Agregar a la lista
            $Results.DriversAnalyzed += $driverInfo

            # Marcar si necesita actualizacion
            if ($needsUpdate) {
                $Results.OutdatedDrivers += $driverInfo
                Write-Host "Driver desactualizado encontrado: $($driver.DeviceName) v$($driver.DriverVersion)" -ForegroundColor Yellow
            }

            # Marcar tarea completada
            $Results.CompletedTasks += [PSCustomObject]@{
                TaskName    = "Analizar driver $($driver.DeviceName)"
                Status      = "Exito"
                Timestamp   = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            }

        }
        catch {
            $errorMsg = "Error procesando driver $($driver.DeviceName): $($_.Exception.Message)"
            Write-Warning $errorMsg

            # Registrar error critico
            $Results.CriticalErrors += [PSCustomObject]@{
                ErrorMessage   = $_.Exception.Message
                ErrorType      = $_.Exception.GetType().Name
                AffectedDriver = $driver.DeviceName
                Timestamp      = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            }
        }
    }

    # Guardar resultados en JSON
    if ($Results.DriversAnalyzed.Count -gt 0) {
        $report = [PSCustomObject]@{
            ScanTimestamp    = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            ComputerName     = $env:COMPUTERNAME
            TotalDrivers     = $Results.DriversAnalyzed.Count
            OutdatedDrivers  = $Results.OutdatedDrivers.Count
            Drivers          = $Results.DriversAnalyzed
        }

        $reportJson = $report | ConvertTo-Json -Depth 5
        $reportJson | Out-File -FilePath $outputPath -Force -Encoding UTF8

        Write-Host "Reporte guardado en: $outputPath" -ForegroundColor Green
        Write-Host "Total drivers analizados: $($Results.DriversAnalyzed.Count)" -ForegroundColor Cyan
        Write-Host "Drivers desactualizados: $($Results.OutdatedDrivers.Count)" -ForegroundColor $(if ($Results.OutdatedDrivers.Count -gt 0) { 'Yellow' } else { 'Green' })
    }

    # Envio de resultados a API REST (si se configura la URL)
    if ($Url -and $Results.DriversAnalyzed.Count -gt 0) {
        try {
            Write-Host "Enviando resultados a API..." -ForegroundColor Cyan
            
            $deviceId = 1  # ID del dispositivo registrado
            $groupId = 1   # ID del grupo en GroupsIntune
            
            # Crear payload con la estructura esperada por la API
            $bodyObject = @{
                id_group = $groupId
                status = if ($Results.OutdatedDrivers.Count -eq 0) { $true } else { $false }
                action_remediation = $Results
                id_device = $deviceId
            }
            
            # Convertir a JSON
            $jsonPayload = $bodyObject | ConvertTo-Json -Depth 10 -Compress
            
            # Guardar payload para depuración
            $debugPath = "$env:TEMP\DriverDetectionPayload.json"
            [System.IO.File]::WriteAllText($debugPath, $jsonPayload, [System.Text.Encoding]::UTF8)
            Write-Host "Payload guardado en: $debugPath" -ForegroundColor Gray
            
            # Enviar a la API
            $response = Invoke-WebRequest -Uri $Url -Method Post -Body $jsonPayload -ContentType "application/json; charset=utf-8" -ErrorAction Stop
            
            Write-Host "Resultados enviados a API exitosamente (Status: $($response.StatusCode))" -ForegroundColor Green
            Write-Host "Respuesta de API: $($response.Content)" -ForegroundColor White
        }
        catch {
            Write-Error "Error enviando resultados a API: $($_.Exception.Message)"
            
            # Capturar detalles del error HTTP
            if ($_.Exception.Response) {
                try {
                    $errorStream = $_.Exception.Response.GetResponseStream()
                    $reader = New-Object System.IO.StreamReader($errorStream)
                    $errorDetails = $reader.ReadToEnd()
                    Write-Host "Detalles del error HTTP:" -ForegroundColor Red
                    Write-Host $errorDetails -ForegroundColor Red
                    $reader.Dispose()
                }
                catch {
                    Write-Host "No se pudieron obtener detalles del error HTTP" -ForegroundColor Yellow
                }
            }
        }
    }

    # Determinar codigo de salida para Intune
    if ($Results.OutdatedDrivers.Count -gt 0) {
        Write-Host "Se requieren actualizaciones de drivers" -ForegroundColor Yellow
        exit 1  # Intune ejecutara el script de remediacion
    }
    else {
        Write-Host "Todos los drivers criticos estan actualizados" -ForegroundColor Green
        exit 0  # No se requiere accion
    }
}
catch {
    Write-Error "Error en la deteccion: $($_.Exception.Message)"
    Write-Host "Tipo de error: $($_.Exception.GetType().Name)" -ForegroundColor Red

    # Crear reporte de error
    $errorReport = [PSCustomObject]@{
        ErrorTimestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        ComputerName   = $env:COMPUTERNAME
        ErrorMessage   = $_.Exception.Message
        ErrorType      = $_.Exception.GetType().Name
        StackTrace     = $_.ScriptStackTrace
    }

    $errorReport | ConvertTo-Json | Out-File -FilePath $outputPath -Force -Encoding UTF8
    exit 1  # Salir con error para que Intune lo registre
}
