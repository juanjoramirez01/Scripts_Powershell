# -----------------------------
# Script de remediación de problemas de disco
# -----------------------------
$Url = "http://PSEGBKPGML3.suramericana.com.co:8080/api/remediations"

# Configuración de inactividad 
[int]$monthsInactive = 6
$cutoffDate = (Get-Date).AddMonths(-$monthsInactive)

# Inicializar colecciones de resultados 
$Results = [PSCustomObject]@{
    CompletedTasks   = @()
    FileLevelErrors  = @()
    CriticalErrors   = @()
}

function Add-Success {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Task
    )
    $script:Results.CompletedTasks += $Task
}

function Add-FileError {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path,
        [Parameter(Mandatory=$true)]
        [string]$Message
    )
    $script:Results.FileLevelErrors += [PSCustomObject]@{
        Item    = $Path
        Error   = $Message
    }
}

function Add-CriticalError {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message
    )
    $script:Results.CriticalErrors += $Message
}

function Clear-Directory {
    param(
        [Parameter(Mandatory=$true)]
        [string]$PathPattern
    )

    try {
        # Check if path exists before attempting to enumerate
        if (-not (Test-Path $PathPattern.Replace('*', '') -ErrorAction SilentlyContinue)) {
            Add-FileError -Path $PathPattern -Message "Path does not exist or is not accessible"
            return
        }
        
        $items = Get-ChildItem -Path $PathPattern -Recurse -Force -ErrorAction Stop
        foreach ($item in $items) {
            try {
                Remove-Item -LiteralPath $item.FullName -Recurse -Force -ErrorAction Stop
                Add-Success "Removed: $($item.FullName)"
            }
            catch {
                Add-FileError -Path $item.FullName -Message $_.Exception.Message
            }
        }
    }
    catch {
        Add-CriticalError "Failed to enumerate path '$PathPattern': $($_.Exception.Message)"
    }
}

# Limpiar Temporales 
try {
    if ($env:SystemRoot) {
        Clear-Directory -Path "$env:SystemRoot\Temp\*"
    } else {
        Add-CriticalError "SystemRoot environment variable not found"
    }
}
catch {
    Add-CriticalError "System temp cleanup error: $($_.Exception.Message)"
}

try {
    if ($env:TEMP) {
        Clear-Directory -Path "$env:TEMP\*"
    } else {
        Add-CriticalError "TEMP environment variable not found"
    }
}
catch {
    Add-CriticalError "User temp cleanup error: $($_.Exception.Message)"
}

# Cache de Navegadores 
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
                Add-CriticalError "Browser cache cleanup error for '$path': $($_.Exception.Message)"
            }
        }
    } else {
        Add-CriticalError "USERPROFILE environment variable not found for browser cleanup"
    }
}
catch {
    Add-CriticalError "Browser cache cleanup initialization error: $($_.Exception.Message)"
}

# Perfiles Inactivos 
try {
    $oldProfiles = Get-CimInstance -Class Win32_UserProfile -ErrorAction Stop | Where-Object {
        -not $_.Special -and 
        $_.Loaded -eq $false -and
        $_.LastUseTime -and
        $_.LastUseTime -ne $null -and
        ([Management.ManagementDateTimeConverter]::ToDateTime($_.LastUseTime) -lt $cutoffDate)
    }

    Add-Success "Detected $($oldProfiles.Count) old profiles"

    foreach ($prof in $oldProfiles) {
        $path = $prof.LocalPath
        
        if ($path -and (Test-Path -Path $path -ErrorAction SilentlyContinue) -and $path -notmatch '^C:\\(Windows|Program Files)') {
            try {
                $prof | Remove-CimInstance -ErrorAction Stop
                
                if (Test-Path -Path $path -ErrorAction SilentlyContinue) {
                    Remove-Item -Path $path -Recurse -Force -ErrorAction Stop
                    Add-Success "Profile removed: $path"
                }
            }
            catch {
                Add-FileError -Path $path -Message $_.Exception.Message
            }
        } else {
            if (-not $path) {
                Add-FileError -Path "Unknown" -Message "Profile path is null or empty"
            } elseif (-not (Test-Path -Path $path -ErrorAction SilentlyContinue)) {
                Add-FileError -Path $path -Message "Profile path does not exist"
            } else {
                Add-FileError -Path $path -Message "Profile path is protected (system directory)"
            }
        }
    }
}
catch {
    Add-CriticalError "Profile detection/removal error: $($_.Exception.Message)"
}

# Salida JSON
try {
    if ($env:LOCALAPPDATA) {
        # Use UTF8NoBOM encoding to prevent Unicode escaping
        $JsonOutput = $Results | ConvertTo-Json -Depth 4
        [System.IO.File]::WriteAllText("$env:LOCALAPPDATA\CleanupResults.json", $JsonOutput, [System.Text.Encoding]::UTF8)
        
        Write-Output $JsonOutput
        
        # Enviar resultados a la API
        try {
            $deviceId = 1
            $groupId = "Auto-created Group: DISK_CLEANUP_20250731"
            
            # Keep Results as an object/array for action_remediation field
            $body = @{
                id_group = $groupId
                status = 1
                action_remediation = $Results
                id_device = $deviceId
            } | ConvertTo-Json -Depth 5
            
            # Log the request details
            # Write-Host "`n=== API REQUEST LOG ===" -ForegroundColor Yellow
            # Write-Host "Timestamp: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Gray
            # Write-Host "Endpoint: $Url" -ForegroundColor Gray
            # Write-Host "Method: POST" -ForegroundColor Gray
            # Write-Host "Headers: Accept=application/json, Content-Type=application/json" -ForegroundColor Gray
            # Write-Host "Request Body Length: $($body.Length) characters" -ForegroundColor Gray
            # Write-Host "Request Body Preview:" -ForegroundColor Gray
            # Write-Host $body -ForegroundColor Cyan
            # Write-Host "========================" -ForegroundColor Yellow
            
            $response = Invoke-WebRequest -Uri $Url -Method POST -Body $body -Headers @{"Accept"="application/json"; "Content-Type"="application/json"} -ErrorAction Stop
            
            # Log the successful response
            # Write-Host "`n=== API RESPONSE LOG ===" -ForegroundColor Green
            # Write-Host "Timestamp: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Gray
            # Write-Host "Status Code: $($response.StatusCode)" -ForegroundColor Green
            # Write-Host "Response Headers:" -ForegroundColor Gray
            # foreach ($header in $response.Headers.Keys) {
            #     Write-Host "  ${header}: $($response.Headers[$header])" -ForegroundColor Gray
            # }
            # Write-Host "Response Body:" -ForegroundColor Gray
            # Write-Host $response.Content -ForegroundColor Cyan
            # Write-Host "=========================" -ForegroundColor Green
            
            Add-Success "API notification sent successfully: $($response.StatusCode)"
            
            # Respuesta de la API
            # Write-Host "Raw Response Content:" -ForegroundColor Gray
            # Write-Host $response.Content -ForegroundColor Gray
            
            try {
                # $responseContent = $response.Content | ConvertFrom-Json
                # Write-Host "Created remediation entry with ID: $($responseContent.data.id)" -ForegroundColor Green
                # Write-Host "Action remediation field contains:" -ForegroundColor Yellow
                # Write-Host ($responseContent.data.action_remediation | ConvertTo-Json -Depth 4) -ForegroundColor Cyan
            } catch {
                # Write-Host "Warning: Could not parse JSON response: $($_.Exception.Message)" -ForegroundColor Yellow
                # Write-Host "Raw response was: $($response.Content)" -ForegroundColor Yellow
                Add-FileError -Path "API Response" -Message "Could not parse JSON response: $($_.Exception.Message)"
            }
            
        } catch {
            $errorMessage = "Failed to send API notification"
            
            # Log the error details
            # Write-Host "`n=== API ERROR LOG ===" -ForegroundColor Red
            # Write-Host "Timestamp: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Gray
            # Write-Host "Error Type: API Request Failed" -ForegroundColor Red
            
            if ($_.Exception.Response) {
                # Write-Host "HTTP Status Code: $($_.Exception.Response.StatusCode)" -ForegroundColor Red
                # Write-Host "HTTP Status Description: $($_.Exception.Response.StatusDescription)" -ForegroundColor Red
                
                try {
                    $errorStream = $_.Exception.Response.GetResponseStream()
                    $reader = New-Object System.IO.StreamReader($errorStream)
                    try {
                        $errorDetails = $reader.ReadToEnd()
                        # Write-Host "Server Response Body:" -ForegroundColor Red
                        # Write-Host $errorDetails -ForegroundColor Yellow
                        
                        # Try to parse as JSON to get better error details
                        try {
                            # $errorJson = $errorDetails | ConvertFrom-Json
                            # Write-Host "Parsed Error Details:" -ForegroundColor Red
                            # Write-Host "Status: $($errorJson.status)" -ForegroundColor Yellow
                            # Write-Host "Message: $($errorJson.message)" -ForegroundColor Yellow
                            # Write-Host "Errors: $($errorJson.errors)" -ForegroundColor Yellow
                        } catch {
                            # Write-Host "Could not parse error response as JSON" -ForegroundColor Yellow
                        }
                        
                        $errorMessage += ": $errorDetails"
                    }
                    finally {
                        $reader.Dispose()
                    }
                } catch {
                    # Write-Host "Could not read error response: $($_.Exception.Message)" -ForegroundColor Red
                    $errorMessage += ": $($_.Exception.Message)"
                }
            } else {
                # Write-Host "Network/Connection Error: $($_.Exception.Message)" -ForegroundColor Red
                # Write-Host "Inner Exception: $($_.Exception.InnerException.Message)" -ForegroundColor Red
                $errorMessage += ": $($_.Exception.Message)"
            }
            
            # Write-Host "Original Request Body:" -ForegroundColor Red
            # Write-Host $body -ForegroundColor Yellow
            # Write-Host "Full Exception Details:" -ForegroundColor Red
            # Write-Host $_.Exception.ToString() -ForegroundColor Gray
            # Write-Host "=====================" -ForegroundColor Red
            
            Add-CriticalError $errorMessage
        }
    } else {
        Add-CriticalError "LOCALAPPDATA environment variable not found"
        $JsonOutput = $Results | ConvertTo-Json -Depth 4
        Write-Output $JsonOutput
        try {
            [System.IO.File]::WriteAllText("CleanupResults.json", $JsonOutput, [System.Text.Encoding]::UTF8)
            # Write-Host "Results saved to CleanupResults.json in the current directory." -ForegroundColor Yellow
        } catch {
            Write-Warning "Could not save results to file: $($_.Exception.Message)"
            Add-CriticalError "Could not save results to file: $($_.Exception.Message)"
        }
    }
}
catch {
    Write-Error "Failed to generate output: $($_.Exception.Message)"
    Add-CriticalError "Failed to generate output: $($_.Exception.Message)"
}

# Final exit logic based on critical errors
if ($Results.CriticalErrors.Count -gt 0) {
    exit 1
    Write-Host "Script completed with critical errors. Check logs for details." -ForegroundColor Red
} else {
    exit 0
    Write-Host "Script completed successfully." -ForegroundColor Green
}