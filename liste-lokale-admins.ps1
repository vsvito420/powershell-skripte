# OU korrekt definieren
$OU = "OU=SBSComputers,OU=computers,OU=myBusiness,DC=geta,DC=local"

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
    
    try {
        # Timeout über Job-Funktionalität implementieren
        $Job = Invoke-Command -ComputerName $Computer -ScriptBlock $ScriptBlock -ErrorAction Stop -AsJob
        
        # Auf den Job warten (mit Timeout von 30 Sekunden)
        $Completed = Wait-Job -Job $Job -Timeout 30
        
        if ($Completed -eq $null) {
            # Timeout ist aufgetreten
            Remove-Job -Job $Job -Force
            throw "Zeitüberschreitung bei der Verbindung"
        }
        
        # Ergebnisse abrufen
        $Admins = Receive-Job -Job $Job
        Remove-Job -Job $Job
        
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
    }
    
    # Kurze Pause zwischen den Computern
    Start-Sleep -Milliseconds 500
}

# Ergebnisse auf dem Desktop speichern
$DesktopPath = [Environment]::GetFolderPath("Desktop")
$ExportPath = "$DesktopPath\LocalAdminsStatus.csv"

$Results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8

Write-Host "`nPrüfung abgeschlossen. Ergebnisse wurden gespeichert unter:" -ForegroundColor Cyan
Write-Host $ExportPath -ForegroundColor White
