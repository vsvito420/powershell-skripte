# OU korrekt definieren
$OU = "OU=SBSComputers,OU=computers,OU=myBusiness,DC=nameOfDomain,DC=local"

# Alle Computer aus der OU abrufen (mit Fehlerbehandlung)
try {
    Write-Host "Rufe Computer aus der OU ab..." -ForegroundColor Cyan
    $Computers = Get-ADComputer -Filter * -SearchBase $OU -ErrorAction Stop | Select-Object -ExpandProperty Name
    Write-Host "Gefunden: $($Computers.Count) Computer" -ForegroundColor Green
}
catch {
    $ErrorMsg = $_.Exception.Message
    Write-Host "Fehler bei der OU-Abfrage: $ErrorMsg" -ForegroundColor Red
    exit
}

# Ergebnisliste initialisieren
$Results = @()

# Skriptblock zur Prüfung der lokalen Admins
$ScriptBlock = {
    Get-LocalGroupMember -Group "Administratoren" | Select-Object Name
}

# Remote-Ausführung - ein Computer nach dem anderen
$Counter = 0
foreach ($Computer in $Computers) {
    $Counter++
    Write-Host "($Counter/$($Computers.Count)) Prüfe Computer: $Computer..." -ForegroundColor Yellow
    
    # Erzwinge die Verwendung von Invoke-Command mit Timeout, ohne vorherigen Ping-Test
    try {
        # Job für den Remoting-Befehl erstellen und starten
        $Job = Invoke-Command -ComputerName $Computer -ScriptBlock $ScriptBlock -ErrorAction Stop -AsJob
        
        # Auf den Job warten (mit Timeout von 3 Sekunden)
        $Timeout = 3
        $CompletedInTime = $Job | Wait-Job -Timeout $Timeout
        
        if (-not $CompletedInTime) {
            # Timeout ist aufgetreten - Job abbrechen
            $Job | Stop-Job -PassThru | Remove-Job -Force
            throw "Zeitüberschreitung bei der Verbindung (${Timeout}s)"
        }
        
        # Ergebnisse abrufen
        $Admins = $Job | Receive-Job -ErrorAction Stop
        $Job | Remove-Job
        
        # Prüfen ob Admins leer ist
        if ($null -eq $Admins -or ($Admins -is [Array] -and $Admins.Count -eq 0)) {
            throw "Keine Administratoren-Daten empfangen"
        }
        
        # Liste der Admins erstellen und anzeigen
        $AdminList = $Admins.Name -join ", "
        Write-Host "  - Erfolgreich: $Computer hat folgende lokale Admins:" -ForegroundColor Green
        Write-Host "    $AdminList" -ForegroundColor White
        
        $Results += [PSCustomObject]@{
            ComputerName = $Computer
            Admins       = $AdminList
            Status       = "Erfolg"
        }
    }
    catch {
        $ErrorMsg = $_.Exception.Message
        Write-Host "  - Fehler bei $Computer $ErrorMsg" -ForegroundColor Red
        
        $Results += [PSCustomObject]@{
            ComputerName = $Computer
            Admins       = "N/A"
            Status       = "Fehler: $ErrorMsg"
        }
        
        # Sicherstellen, dass keine Jobs zurückbleiben
        Get-Job | Where-Object { $_.Name -match $Computer -or $_.Location -match $Computer } | 
        Remove-Job -Force -ErrorAction SilentlyContinue
    }
    finally {
        # Genereller Cleanup für verwaiste Jobs
        Get-Job | Where-Object { $_.State -eq 'Running' -and $_.PSBeginTime -lt (Get-Date).AddSeconds(-10) } | 
        Stop-Job -PassThru | Remove-Job -Force -ErrorAction SilentlyContinue
    }
    
    # Kurze Pause zwischen den Computern
    Start-Sleep -Milliseconds 100
}

# Ergebnisse auf dem Desktop speichern
$DesktopPath = [Environment]::GetFolderPath("Desktop")
$ExportPath = "$DesktopPath\LocalAdminsStatus.csv"
$Results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8

# Auch eine HTML-Version für bessere Lesbarkeit erstellen
$HtmlPath = "$DesktopPath\LocalAdminsStatus.html"
$HtmlHead = @"
<style>
    table { border-collapse: collapse; width: 100%; }
    th, td { text-align: left; padding: 8px; border-bottom: 1px solid #ddd; }
    tr:nth-child(even) { background-color: #f2f2f2; }
    th { background-color: #4CAF50; color: white; }
    .error { color: red; }
    .success { color: green; }
</style>
"@

$Results | ConvertTo-Html -Head $HtmlHead -Property ComputerName, Admins, Status | 
ForEach-Object { 
    $_ -replace '<td>Fehler:', '<td class="error">Fehler:' -replace '<td>Erfolg</td>', '<td class="success">Erfolg</td>' 
} | 
Out-File -FilePath $HtmlPath -Encoding utf8

Write-Host "`nPrüfung abgeschlossen. Ergebnisse wurden gespeichert unter:" -ForegroundColor Cyan
Write-Host "CSV: $ExportPath" -ForegroundColor White
Write-Host "HTML: $HtmlPath" -ForegroundColor White

# Statistik anzeigen
$Successful = ($Results | Where-Object { $_.Status -eq "Erfolg" }).Count
$Failed = $Computers.Count - $Successful
Write-Host "`nZusammenfassung:" -ForegroundColor Cyan
Write-Host "Erfolgreich: $Successful" -ForegroundColor Green
Write-Host "Fehlgeschlagen: $Failed" -ForegroundColor $(if ($Failed -gt 0) { "Red" } else { "Green" })
Write-Host "Erfolgsrate: $([Math]::Round(($Successful / $Computers.Count) * 100, 1))%" -ForegroundColor Cyan
