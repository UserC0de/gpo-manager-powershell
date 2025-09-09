<#
.SYNOPSIS
  Gestor de GPOs interactivo endurecido

.DESCRIPTION
  Menu para listar, buscar, respaldar, reportar enlaces y eliminar GPOs con prácticas de producción:
  - Set-StrictMode y errores terminantes
  - Parsing XML para LinksTo/SOMPath
  - Backups con metadatos y carpeta por ejecución
  - Confirmaciones seguras e integración -WhatIf/-Confirm

.REQUIREMENTS
  - Windows con RSAT/GPMC (módulo GroupPolicy)
#>

# Endurecer ejecución
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$PSDefaultParameterValues['*:ErrorAction'] = 'Stop'

# Verbose por defecto controlable via $VerbosePreference
$VerbosePreference = 'Continue'

# Utilidades
function New-SafeDirectory {
    param(
        [Parameter(Mandatory)][string]$Path
    )
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -Path $Path -ItemType Directory | Out-Null
    }
}

function Get-GpoLinks {
    param(
        [Parameter(Mandatory)][Guid]$GpoId
    )
    # Devuelve array de SOMPath; vacío si no hay enlaces
    [xml]$xml = Get-GPOReport -Guid $GpoId -ReportType Xml
    # Algunas GPOs pueden no tener LinksTo; proteger acceso
    $linksNode = $xml.GPO.LinksTo
    if ($null -eq $linksNode) { return @() }
    # LinksTo puede contener múltiples SOMPath
    $som = @()
    foreach ($p in $linksNode.SOMPath) {
        if ($p -and $p.'#text') { $som += $p.'#text' } else { $som += [string]$p }
    }
    # Filtrar vacíos
    $som | Where-Object { $_ -and $_.Trim().Length -gt 0 }
}

function Mostrar-Menu {
    Clear-Host
    Write-Host "========================="
    Write-Host "GESTOR DE GPOs"
    Write-Host "========================="
    Write-Host "1. Ver todas las GPOs"
    Write-Host "2. Ver enlaces de una GPO"
    Write-Host "3. Backup de una GPO"
    Write-Host "4. Eliminar una GPO (con confirmación)"
    Write-Host "5. Backup de TODAS las GPOs"
    Write-Host "6. Buscar GPO por nombre parcial"
    Write-Host "7. Ver enlaces activos de TODAS las GPOs"
    Write-Host "8. Eliminar GPOs sin enlaces"
    Write-Host "9. Salir"
    Write-Host "========================="
}

function Listar-GPOs {
    $gpos = Get-GPO -All | Select-Object DisplayName, Id
    $gpos | Sort-Object DisplayName | Format-Table -AutoSize
    Pause
}

function Ver-Enlace-GPO {
    $nombre = Read-Host "Nombre de la GPO"
    if (-not $nombre) { Write-Warning "Nombre vacío"; Pause; return }
    try {
        $gpo = Get-GPO -Name $nombre
        $informe = Join-Path $env:USERPROFILE "Desktop\GPO_$($gpo.DisplayName -replace '[^\w.-]', '_').html"
        Get-GPOReport -Guid $gpo.Id -ReportType Html -Path $informe
        $links = Get-GpoLinks -GpoId $gpo.Id
        Write-Host "Enlaces (SOMPath):"
        if ($links.Count -eq 0) { Write-Host " - Sin enlaces" } else { $links | ForEach-Object { Write-Host " - $_" } }
        Start-Process $informe
    } catch {
        Write-Warning "No se pudo generar u abrir el informe para '$nombre' - $($_.Exception.Message)"
    }
    Pause
}

function Backup-UnaGPO {
    $nombre = Read-Host "Nombre de la GPO"
    if (-not $nombre) { Write-Warning "Nombre vacío"; Pause; return }
    try {
        $gpo = Get-GPO -Name $nombre
        $timestamp = Get-Date -Format yyyyMMdd_HHmmss
        $base = "C:\BackupGPOs"
        New-SafeDirectory -Path $base
        $ruta = Join-Path $base "$($gpo.DisplayName -replace '[^\w.-]', '_')-$timestamp"
        New-SafeDirectory -Path $ruta
        Backup-GPO -Guid $gpo.Id -Path $ruta -Comment "Auto-backup $timestamp"
        # Guardar reporte XML con metadatos (para cubrir enlaces/WMI filters fuera del backup)
        $xmlPath = Join-Path $ruta "GPOReport.xml"
        Get-GPOReport -Guid $gpo.Id -ReportType Xml -Path $xmlPath
        Write-Host "Backup guardado en: $ruta"
    } catch {
        Write-Warning "Error al hacer backup - $($_.Exception.Message)"
    }
    Pause
}

function Backup-TodasLasGPOs {
    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $rutaBase = "C:\BackupGPOs\$timestamp"
    New-SafeDirectory -Path $rutaBase
    $all = Get-GPO -All
    foreach ($gpo in $all) {
        try {
            Backup-GPO -Guid $gpo.Id -Path $rutaBase -Comment "Auto-backup $timestamp"
            $xmlPath = Join-Path $rutaBase "$($gpo.DisplayName -replace '[^\w.-]', '_').xml"
            Get-GPOReport -Guid $gpo.Id -ReportType Xml -Path $xmlPath
            Write-Verbose "Backup correcto de: $($gpo.DisplayName)"
        } catch {
            Write-Warning "Error al respaldar '$($gpo.DisplayName)' - $($_.Exception.Message)"
        }
    }
    Write-Host "Backup completo guardado en: $rutaBase"
    Pause
}

function Buscar-GPO {
    $filtro = Read-Host "Introduce parte del nombre de la GPO"
    if (-not $filtro) { Write-Warning "Filtro vacío"; Pause; return }
    Get-GPO -All |
        Where-Object { $_.DisplayName -like "*$filtro*" } |
        Select-Object DisplayName, Id |
        Sort-Object DisplayName |
        Format-Table -AutoSize
    Pause
}

function Ver-EnlacesTodasGPOs {
    try {
        $informe = Join-Path $env:USERPROFILE "Desktop\GPO_Todas_Links.html"
        Get-GPO -All | Get-GPOReport -ReportType Html -Path $informe
        # Resumen en consola consultando XML por GPO
        $all = Get-GPO -All
$all = Get-GPO -All
$(foreach ($gpo in $all) {
    $links = Get-GpoLinks -GpoId $gpo.Id
    $joined = if ($links.Count) { $links -join '; ' } else { 'Sin enlaces' }
    [pscustomobject]@{ Name = $gpo.DisplayName; Links = $joined }
}) | Sort-Object Name | Format-Table -AutoSize
        Start-Process $informe
    } catch {
        Write-Warning "No se pudo generar el informe de enlaces - $($_.Exception.Message)"
    }
    Pause
}

function Eliminar-GPOsSinEnlace {
    Write-Host "Buscando GPOs sin enlaces..."
    $todasGPOs = Get-GPO -All
    $huerfanas = foreach ($gpo in $todasGPOs) {
        $links = Get-GpoLinks -GpoId $gpo.Id
        if ($links.Count -eq 0) { $gpo }
    }
    if (-not $huerfanas -or $huerfanas.Count -eq 0) {
        Write-Host "No se encontraron GPOs huérfanas."
    } else {
        Write-Host "Se encontraron $($huerfanas.Count) GPOs sin enlace."
        foreach ($gpo in $huerfanas) {
            $confirm = Read-Host "Eliminar '$($gpo.DisplayName)'? (s/n)"
            if ($confirm -eq "s") {
                try {
                    # Opción: Backup automático previo
                    $doBackup = Read-Host "¿Hacer backup previo? (s/n)"
                    if ($doBackup -eq 's') {
                        $timestamp = Get-Date -Format yyyyMMdd_HHmmss
                        $ruta = "C:\BackupGPOs\PreDelete-$timestamp\$($gpo.DisplayName -replace '[^\w.-]', '_')"
                        New-SafeDirectory -Path (Split-Path -Parent $ruta)
                        New-SafeDirectory -Path $ruta
                        Backup-GPO -Guid $gpo.Id -Path $ruta -Comment "Pre-delete backup $timestamp"
                        Get-GPOReport -Guid $gpo.Id -ReportType Xml -Path (Join-Path $ruta "GPOReport.xml")
                    }
                    Remove-GPO -Guid $gpo.Id -Confirm:$true
                    Write-Host "GPO eliminada: $($gpo.DisplayName)"
                } catch {
                    Write-Warning "No se pudo eliminar '$($gpo.DisplayName)' - $($_.Exception.Message)"
                }
            }
        }
    }
    Pause
}

function Eliminar-GPO {
    $nombre = Read-Host "Nombre de la GPO a eliminar"
    if (-not $nombre) { Write-Warning "Nombre vacío"; Pause; return }
    try {
        $gpo = Get-GPO -Name $nombre
    } catch {
        Write-Warning "GPO no encontrada: $nombre"; Pause; return
    }
    $confirm = Read-Host "¿Seguro que quieres eliminar '$($gpo.DisplayName)'? (s/n)"
    if ($confirm -ne "s") { Write-Host "Operación cancelada."; Pause; return }

    $backup = Read-Host "¿Quieres hacer un backup antes? (s/n)"
    if ($backup -eq "s") {
        $timestamp = Get-Date -Format yyyyMMdd_HHmmss
        $ruta = "C:\BackupGPOs\PreDelete-$timestamp\$($gpo.DisplayName -replace '[^\w.-]', '_')"
        New-SafeDirectory -Path (Split-Path -Parent $ruta)
        New-SafeDirectory -Path $ruta
        try {
            Backup-GPO -Guid $gpo.Id -Path $ruta -Comment "Pre-delete backup $timestamp"
            Get-GPOReport -Guid $gpo.Id -ReportType Xml -Path (Join-Path $ruta "GPOReport.xml")
            Write-Host "Backup hecho en: $ruta"
        } catch {
            Write-Warning "No se pudo hacer el backup - $($_.Exception.Message)"
            Pause; return
        }
    }
    try {
        Remove-GPO -Guid $gpo.Id -Confirm:$true
        Write-Host "GPO eliminada correctamente"
    } catch {
        Write-Warning "Error al eliminar la GPO - $($_.Exception.Message)"
    }
    Pause
}

# Bucle principal del menú
Do {
    Mostrar-Menu
    $opcion = Read-Host "Selecciona una opción"
    switch ($opcion) {
        "1" { Listar-GPOs }
        "2" { Ver-Enlace-GPO }
        "3" { Backup-UnaGPO }
        "4" { Eliminar-GPO }
        "5" { Backup-TodasLasGPOs }
        "6" { Buscar-GPO }
        "7" { Ver-EnlacesTodasGPOs }
        "8" { Eliminar-GPOsSinEnlace }
        "9" { Write-Host "Saliendo..."; exit }
        default { Write-Host "Opción no válida"; Pause }
    }
} while ($true)
