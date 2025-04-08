# Skript zum Entfernen aller Benutzer aus der Administratorengruppe außer dem Domänen-Administrator
# Dieses Skript ist für die Ausführung über das Ninja RMM Tool konzipiert

try {
    # Verzeichnis erstellen, falls es nicht existiert
    $logDir = "C:\afi-temp"
    if (-not (Test-Path -Path $logDir)) {
        New-Item -Path $logDir -ItemType Directory -Force | Out-Null
        Write-Host "Verzeichnis $logDir wurde erstellt."
    }
    
    # Logdatei im afi-temp Verzeichnis erstellen oder anhängen
    $logPath = "$logDir\admin-cleanup-log.txt"
    Start-Transcript -Path $logPath -Append -Force
    
    Write-Host "Start des Skripts: $(Get-Date)"
    
    # Bestimmen, ob Computer in einer Domäne ist
    $inDomain = (Get-WmiObject -Class Win32_ComputerSystem).PartOfDomain
    Write-Host "Computer ist Teil einer Domäne: $inDomain"
    
    # Administratorengruppe abrufen - Berücksichtigung von lokalisierten Namen
    if ([System.Environment]::OSVersion.Version.Major -ge 6) {
        # Für Windows Vista oder höher, SID für die Administratorengruppe ist verlässlicher
        $adminGroup = Get-LocalGroup -SID "S-1-5-32-544"
    } else {
        # Für ältere Betriebssysteme
        $adminGroup = [ADSI]"WinNT://./Administrators,group"
    }
    
    Write-Host "Administratorengruppe gefunden: $($adminGroup.Name)"
    
    # Mitglieder der Administratorengruppe abrufen
    $members = Get-LocalGroupMember -Group $adminGroup.Name
    Write-Host "Gefundene Mitglieder in der Administratorengruppe: $($members.Count)"
    
    # Domänen-Admin-Konto identifizieren
    $domainAdminName = if ($inDomain) { "Domain Admins" } else { "Administrator" }
    Write-Host "Zu behaltender Admin-Account: $domainAdminName"
    
    # Jeden Benutzer überprüfen und gegebenenfalls entfernen
    foreach ($member in $members) {
        $memberName = $member.Name
        # Name nach dem Backslash extrahieren, falls vorhanden (Domain\Username)
        if ($memberName -match '\\') {
            $shortName = $memberName.Split('\')[1]
        } else {
            $shortName = $memberName
        }
        
        Write-Host "Überprüfe Mitglied: $memberName (Kurzname: $shortName)"
        
        # Prüfen, ob der Benutzer behalten werden soll
        if (($shortName -ne $domainAdminName) -and ($shortName -ne "Administrator") -and ($shortName -notlike "*Domain Admins*")) {
            try {
                Remove-LocalGroupMember -Group $adminGroup.Name -Member $memberName
                Write-Host "Benutzer $memberName wurde aus der Administratorengruppe entfernt."
            } catch {
                Write-Host "Fehler beim Entfernen von $memberName : $_" -ForegroundColor Red
            }
        } else {
            Write-Host "Behalte $memberName als Administrator." -ForegroundColor Green
        }
    }
    
    Write-Host "Skript erfolgreich abgeschlossen: $(Get-Date)"
} catch {
    Write-Host "Ein Fehler ist aufgetreten: $_" -ForegroundColor Red
} finally {
    Stop-Transcript
    
    # Ausgabe für Ninja RMM
    Write-Host "Log wurde erstellt unter: $logPath"
    Write-Host "#NINJA-CUSTOM-EXIT-CODE:0"
}
