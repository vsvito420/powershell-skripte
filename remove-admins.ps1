# Skript zum Entfernen aller Benutzer aus der Administratorengruppe außer dem Domänen-Administrator
# Dieses Skript ist für die Ausführung über das Ninja RMM Tool konzipiert
# Unterstützt deutsche und englische Windows-Installationen

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
    
    # Administratorengruppe mittels SID abrufen - funktioniert unabhängig von der Sprache
    try {
        # Für Windows Vista oder höher (PowerShell 5.1+)
        $adminGroup = Get-LocalGroup -SID "S-1-5-32-544"
        Write-Host "Administratorengruppe gefunden über SID: $($adminGroup.Name)"
    } catch {
        # Alternativer Ansatz für ältere/andere Systeme
        $adminGroupNames = @("Administratoren", "Administrators")
        $foundGroup = $false
        
        foreach ($groupName in $adminGroupNames) {
            try {
                $adminGroup = Get-LocalGroup -Name $groupName -ErrorAction Stop
                Write-Host "Administratorengruppe gefunden mit Name: $groupName"
                $foundGroup = $true
                break
            } catch {
                Write-Host "Gruppe '$groupName' nicht gefunden, versuche alternative Namen..."
            }
        }
        
        if (-not $foundGroup) {
            throw "Konnte die Administratorengruppe nicht finden. Bitte überprüfen Sie die Gruppenbezeichnungen."
        }
    }
    
    # Mitglieder der Administratorengruppe abrufen
    $members = Get-LocalGroupMember -Group $adminGroup.Name
    Write-Host "Gefundene Mitglieder in der Administratorengruppe: $($members.Count)"
    
    # Domänen-Admin-Konto identifizieren (mehrsprachig)
    if ($inDomain) {
        $domainAdminNames = @("Domain Admins", "Domänen-Admins", "Domänenadministratoren")
    } else {
        $domainAdminNames = @("Administrator", "Admini strator")
    }
    
    Write-Host "Zu behaltende Admin-Accounts: $($domainAdminNames -join ', ')"
    
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
        $keepUser = $false
        
        # Prüfen gegen bekannte Admin-Namen
        foreach ($adminName in $domainAdminNames) {
            if ($shortName -eq $adminName -or $shortName -like "*$adminName*") {
                $keepUser = $true
                break
            }
        }
        
        # Built-in Administrator wird immer behalten
        if ($member.SID.Value.EndsWith("-500")) {
            Write-Host "Mitglied mit SID $($member.SID.Value) ist der integrierte Administrator-Account."
            $keepUser = $true
        }
        
        if (-not $keepUser) {
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
