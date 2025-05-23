function Start-PrinterDiagnose {
    Clear-Host
    Write-Host "=== DRUCKER DIAGNOSE TOOL ===" -ForegroundColor Cyan
    Write-Host "==============================" -ForegroundColor Cyan
    
    # Alle Drucker abrufen
    Write-Host "`nSuche nach installierten Druckern..." -ForegroundColor Yellow
    
    try {
        $printers = Get-Printer | Sort-Object Name
        
        if ($printers.Count -eq 0) {
            Write-Host "Keine Drucker gefunden!" -ForegroundColor Red
            return
        }
        
        # Drucker-Liste anzeigen
        Write-Host "`nGefundene Drucker:" -ForegroundColor Green
        Write-Host "==================" -ForegroundColor Green
        
        for ($i = 0; $i -lt $printers.Count; $i++) {
            $status = if ($printers[$i].PrinterStatus -eq "Normal") { "✓" } else { "⚠" }
            $gpoIndicator = if (Test-PrinterInGPO -PrinterName $printers[$i].Name) { "[GPO]" } else { "" }
            Write-Host "$($i + 1). $status $($printers[$i].Name) $gpoIndicator" -ForegroundColor White
            Write-Host "    Typ: $($printers[$i].DeviceType) | Status: $($printers[$i].PrinterStatus)" -ForegroundColor Gray
        }
        
        # Benutzer-Auswahl
        Write-Host "`nOptionen:" -ForegroundColor Cyan
        Write-Host "- Geben Sie eine Nummer (1-$($printers.Count)) ein, um einen Drucker auszuwählen"
        Write-Host "- Geben Sie 'N' ein, um einen Druckernamen manuell einzugeben"
        Write-Host "- Geben Sie 'A' ein, um alle GPO-Drucker anzuzeigen"
        Write-Host "- Geben Sie 'Q' ein, um zu beenden"
        
        do {
            $choice = Read-Host "`nIhre Auswahl"
            
            # Beenden
            if ($choice -eq 'Q' -or $choice -eq 'q') {
                Write-Host "Programm beendet." -ForegroundColor Yellow
                return
            }
            
            # Alle GPO-Drucker anzeigen
            if ($choice -eq 'A' -or $choice -eq 'a') {
                Show-AllGPOPrinters
                continue
            }
            
            # Manueller Name
            if ($choice -eq 'N' -or $choice -eq 'n') {
                $printerName = Read-Host "Geben Sie den Druckernamen ein"
                if ([string]::IsNullOrWhiteSpace($printerName)) {
                    Write-Host "Ungültiger Druckername!" -ForegroundColor Red
                    continue
                }
                Start-PrinterDiagnostics -PrinterName $printerName
                return
            }
            
            # Nummer-Auswahl
            if ($choice -match '^\d+$') {
                $index = [int]$choice - 1
                if ($index -ge 0 -and $index -lt $printers.Count) {
                    $selectedPrinter = $printers[$index].Name
                    Start-PrinterDiagnostics -PrinterName $selectedPrinter
                    return
                }
                else {
                    Write-Host "Ungültige Nummer! Bitte wählen Sie zwischen 1 und $($printers.Count)." -ForegroundColor Red
                }
            }
            else {
                Write-Host "Ungültige Eingabe! Bitte geben Sie eine Nummer, 'A', 'N' oder 'Q' ein." -ForegroundColor Red
            }
        } while ($true)
        
    }
    catch {
        Write-Host "Fehler beim Abrufen der Drucker: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Test-PrinterInGPO {
    param([string]$PrinterName)
    
    try {
        # Schnelle Prüfung ob Drucker in GPO-Registry gefunden wird
        $gpoRegPaths = @(
            "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Printers\$PrinterName\GPO*",
            "HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Printers\*",
            "HKCU:\SOFTWARE\Policies\Microsoft\Windows NT\Printers\*"
        )
        
        foreach ($path in $gpoRegPaths) {
            if (Get-ChildItem -Path $path -ErrorAction SilentlyContinue) {
                return $true
            }
        }
        
        # Prüfung in Gruppenrichtlinien-Events
        $events = Get-WinEvent -FilterHashtable @{LogName = 'Microsoft-Windows-GroupPolicy/Operational'; ID = 4016, 5016 } -MaxEvents 50 -ErrorAction SilentlyContinue
        if ($events) {
            foreach ($event in $events) {
                if ($event.Message -like "*$PrinterName*" -or $event.Message -like "*printer*") {
                    return $true
                }
            }
        }
        
        return $false
    }
    catch {
        return $false
    }
}

function Get-PrinterGPOInformation {
    param([string]$PrinterName)
    
    $gpoInfo = @{
        GPOPrinters     = @()
        RegistryEntries = @()
        Events          = @()
        Policies        = @()
    }
    
    try {
        Write-Host "`n8. GRUPPENRICHTLINIEN (GPO) INFORMATIONEN" -ForegroundColor Cyan
        Write-Host "=========================================" -ForegroundColor Cyan
        
        # 1. Registry-basierte GPO-Einträge prüfen
        Write-Host "`n8.1 Registry-Einträge:" -ForegroundColor Yellow
        
        $registryPaths = @(
            @{Path = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Printers\$PrinterName"; Description = "Drucker-spezifische Einstellungen" },
            @{Path = "HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Printers"; Description = "Computer-Drucker-Richtlinien" },
            @{Path = "HKCU:\SOFTWARE\Policies\Microsoft\Windows NT\Printers"; Description = "Benutzer-Drucker-Richtlinien" },
            @{Path = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Connections"; Description = "Drucker-Verbindungen" },
            @{Path = "HKCU:\Printers\Connections"; Description = "Benutzer-Drucker-Verbindungen" }
        )
        
        $foundRegistryEntries = $false
        foreach ($regPath in $registryPaths) {
            try {
                if (Test-Path $regPath.Path) {
                    $items = Get-ChildItem -Path $regPath.Path -ErrorAction SilentlyContinue | 
                    Where-Object { $_.Name -like "*$PrinterName*" -or $_.PSChildName -like "*$PrinterName*" }
                    
                    if ($items) {
                        $foundRegistryEntries = $true
                        Write-Host "  ✓ $($regPath.Description):" -ForegroundColor Green
                        foreach ($item in $items) {
                            Write-Host "    - $($item.PSChildName)" -ForegroundColor Gray
                            $gpoInfo.RegistryEntries += $item
                        }
                    }
                    
                    # Prüfung auf Printer-spezifische Werte
                    $values = Get-ItemProperty -Path $regPath.Path -ErrorAction SilentlyContinue
                    if ($values) {
                        $printerValues = $values.PSObject.Properties | Where-Object { $_.Name -like "*$PrinterName*" -or $_.Value -like "*$PrinterName*" }
                        if ($printerValues) {
                            $foundRegistryEntries = $true
                            Write-Host "  ✓ Gefundene Werte in $($regPath.Description):" -ForegroundColor Green
                            foreach ($value in $printerValues) {
                                Write-Host "    - $($value.Name): $($value.Value)" -ForegroundColor Gray
                            }
                        }
                    }
                }
            }
            catch {
                # Stumm ignorieren wenn Pfad nicht zugänglich
            }
        }
        
        if (-not $foundRegistryEntries) {
            Write-Host "  ℹ Keine druckerspezifischen Registry-Einträge gefunden" -ForegroundColor Gray
        }
        
        # 2. Gruppenrichtlinien-Events prüfen
        Write-Host "`n8.2 Gruppenrichtlinien-Events:" -ForegroundColor Yellow
        
        try {
            $gpoEvents = Get-WinEvent -FilterHashtable @{
                LogName = 'Microsoft-Windows-GroupPolicy/Operational'
                ID      = 4016, 5016, 7016  # GPO-Anwendung Events
            } -MaxEvents 100 -ErrorAction SilentlyContinue | 
            Where-Object { $_.Message -like "*printer*" -or $_.Message -like "*$PrinterName*" }
            
            if ($gpoEvents) {
                Write-Host "  ✓ Gefundene Ereignisse:" -ForegroundColor Green
                foreach ($event in $gpoEvents | Select-Object -First 5) {
                    Write-Host "    - $($event.TimeCreated): ID $($event.Id)" -ForegroundColor Gray
                    Write-Host "      $($event.Message.Substring(0, [Math]::Min(100, $event.Message.Length)))..." -ForegroundColor DarkGray
                    $gpoInfo.Events += $event
                }
                if ($gpoEvents.Count -gt 5) {
                    Write-Host "    ... und $($gpoEvents.Count - 5) weitere Events" -ForegroundColor DarkGray
                }
            }
            else {
                Write-Host "  ℹ Keine relevanten Gruppenrichtlinien-Events gefunden" -ForegroundColor Gray
            }
        }
        catch {
            Write-Host "  ⚠ Fehler beim Abrufen der Events: $($_.Exception.Message)" -ForegroundColor Yellow
        }
        
        # 3. Deployed Printers via GPO
        Write-Host "`n8.3 GPO-deployed Drucker:" -ForegroundColor Yellow
        
        try {
            # WMI-Abfrage für bereitgestellte Drucker
            $deployedPrinters = Get-WmiObject -Class Win32_Printer | 
            Where-Object { $_.Name -eq $PrinterName -and ($_.Attributes -band 0x00000010) }
            
            if ($deployedPrinters) {
                $gpoInfo.GPOPrinters += $deployedPrinters
                Write-Host "  ✓ Drucker via GPO bereitgestellt:" -ForegroundColor Green
                foreach ($printer in $deployedPrinters) {
                    Write-Host "    - Name: $($printer.Name)" -ForegroundColor Gray
                    Write-Host "    - Port: $($printer.PortName)" -ForegroundColor Gray
                    Write-Host "    - Status: $($printer.PrinterStatus)" -ForegroundColor Gray
                }
            }
            else {
                Write-Host "  ℹ Drucker wurde nicht via GPO bereitgestellt" -ForegroundColor Gray
            }
        }
        catch {
            Write-Host "  ⚠ Fehler bei GPO-Drucker-Abfrage: $($_.Exception.Message)" -ForegroundColor Yellow
        }
        
        # 4. Aktuelle Gruppenrichtlinien-Einstellungen
        Write-Host "`n8.4 Aktuelle Richtlinien-Einstellungen:" -ForegroundColor Yellow
        
        try {
            # RSOP (Resultant Set of Policy) Informationen
            if (Get-Command Get-GPResultantSetOfPolicy -ErrorAction SilentlyContinue) {
                Write-Host "  ℹ RSOP-Unterstützung verfügbar - verwenden Sie 'gpresult /h report.html' für detaillierte Informationen" -ForegroundColor Gray
            }
            
            # Letzte GPO-Aktualisierung
            $lastGPUpdate = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\History" -ErrorAction SilentlyContinue
            if ($lastGPUpdate) {
                Write-Host "  ✓ Letzte GPO-Aktualisierung gefunden" -ForegroundColor Green
            }
        }
        catch {
            Write-Host "  ℹ Keine zusätzlichen Richtlinien-Informationen verfügbar" -ForegroundColor Gray
        }
        
        # 5. Empfohlene Aktionen
        if ($foundRegistryEntries -or $gpoEvents -or $deployedPrinters) {
            Write-Host "`n8.5 Empfohlene Aktionen:" -ForegroundColor Yellow
            Write-Host "  • Führen Sie 'gpupdate /force' aus, um GPO zu aktualisieren" -ForegroundColor Cyan
            Write-Host "  • Prüfen Sie 'gpresult /r' für aktuelle Richtlinien" -ForegroundColor Cyan
            Write-Host "  • Verwenden Sie die GPO-Management-Konsole für Details" -ForegroundColor Cyan
        }
        
    }
    catch {
        Write-Host "Fehler bei der GPO-Analyse: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    return $gpoInfo
}

function Show-AllGPOPrinters {
    Clear-Host
    Write-Host "=== ALLE GPO-DRUCKER ÜBERSICHT ===" -ForegroundColor Magenta
    Write-Host "==================================" -ForegroundColor Magenta
    
    try {
        # Alle Drucker abrufen
        $allPrinters = Get-Printer
        
        Write-Host "`nSuche nach GPO-verwalteten Druckern..." -ForegroundColor Yellow
        
        $gpoPrinters = @()
        foreach ($printer in $allPrinters) {
            if (Test-PrinterInGPO -PrinterName $printer.Name) {
                $gpoPrinters += $printer
            }
        }
        
        if ($gpoPrinters.Count -eq 0) {
            Write-Host "Keine GPO-verwalteten Drucker gefunden." -ForegroundColor Yellow
        }
        else {
            Write-Host "`nGefundene GPO-Drucker ($($gpoPrinters.Count)):" -ForegroundColor Green
            Write-Host "=================================" -ForegroundColor Green
            
            foreach ($printer in $gpoPrinters) {
                Write-Host "`n📄 $($printer.Name)" -ForegroundColor White
                Write-Host "   Status: $($printer.PrinterStatus)" -ForegroundColor Gray
                Write-Host "   Typ: $($printer.DeviceType)" -ForegroundColor Gray
                Write-Host "   Treiber: $($printer.DriverName)" -ForegroundColor Gray
                
                # Kurze GPO-Info
                $quickGPO = Get-SimpleGPOInfo -PrinterName $printer.Name
                if ($quickGPO) {
                    Write-Host "   GPO: $quickGPO" -ForegroundColor Cyan
                }
            }
        }
    }
    catch {
        Write-Host "Fehler beim Abrufen der GPO-Drucker: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Write-Host "`nDrücken Sie eine beliebige Taste, um zurückzukehren..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

function Get-SimpleGPOInfo {
    param([string]$PrinterName)
    
    try {
        # Vereinfachte GPO-Info für Übersicht
        $gpoTypes = @()
        
        if (Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Printers\*$PrinterName*" -ErrorAction SilentlyContinue) {
            $gpoTypes += "Computer-Policy"
        }
        
        if (Get-ItemProperty -Path "HKCU:\SOFTWARE\Policies\Microsoft\Windows NT\Printers\*$PrinterName*" -ErrorAction SilentlyContinue) {
            $gpoTypes += "User-Policy"
        }
        
        $deployedPrinter = Get-WmiObject -Class Win32_Printer | 
        Where-Object { $_.Name -eq $PrinterName -and ($_.Attributes -band 0x00000010) }
        if ($deployedPrinter) {
            $gpoTypes += "Deployed"
        }
        
        return ($gpoTypes -join ", ")
    }
    catch {
        return $null
    }
}

function Start-PrinterDiagnostics {
    param([string]$PrinterName)
    
    Clear-Host
    Write-Host "=== DIAGNOSE FÜR DRUCKER: $PrinterName ===" -ForegroundColor Green
    Write-Host "=" * (30 + $PrinterName.Length) -ForegroundColor Green
    
    try {
        # Drucker existiert prüfen
        $printer = Get-Printer -Name $PrinterName -ErrorAction Stop
        
        Write-Host "`n1. GRUNDINFORMATIONEN" -ForegroundColor Cyan
        Write-Host "=====================" -ForegroundColor Cyan
        $printer | Select-Object Name, PrinterStatus, DeviceType, DriverName, Location, Comment | Format-List
        
        Write-Host "`n2. DETAILLIERTER STATUS" -ForegroundColor Cyan
        Write-Host "=======================" -ForegroundColor Cyan
        $wmiPrinter = Get-WmiObject -Class Win32_Printer -Filter "Name='$($PrinterName.Replace("'", "''"))'"
        if ($wmiPrinter) {
            Write-Host "Offline: $(if($wmiPrinter.WorkOffline){'Ja'}else{'Nein'})" -ForegroundColor $(if ($wmiPrinter.WorkOffline) { 'Red' }else { 'Green' })
            Write-Host "Status-Code: $($wmiPrinter.PrinterState)"
            Write-Host "Fehler-Status: $(if($wmiPrinter.DetectedErrorState -eq 0){'Kein Fehler'}else{$wmiPrinter.DetectedErrorState})" -ForegroundColor $(if ($wmiPrinter.DetectedErrorState -eq 0) { 'Green' }else { 'Red' })
        }
        
        Write-Host "`n3. DRUCKAUFTRÄGE" -ForegroundColor Cyan
        Write-Host "================" -ForegroundColor Cyan
        $printJobs = Get-PrintJob -PrinterName $PrinterName -ErrorAction SilentlyContinue
        if ($printJobs) {
            $printJobs | Select-Object DocumentName, UserName, Size, JobStatus, SubmittedTime | Format-Table -AutoSize
            Write-Host "Anzahl Druckaufträge: $($printJobs.Count)" -ForegroundColor Yellow
        }
        else {
            Write-Host "✓ Keine Druckaufträge in der Warteschlange" -ForegroundColor Green
        }
        
        Write-Host "`n4. KONFIGURATION" -ForegroundColor Cyan
        Write-Host "================" -ForegroundColor Cyan
        $config = Get-PrintConfiguration -PrinterName $PrinterName -ErrorAction SilentlyContinue
        if ($config) {
            $config | Select-Object DuplexingMode, PaperSize, Color | Format-List
        }
        
        Write-Host "`n5. PORT-INFORMATIONEN" -ForegroundColor Cyan
        Write-Host "=====================" -ForegroundColor Cyan
        $port = Get-PrinterPort | Where-Object { $_.Name -eq $printer.PortName }
        if ($port) {
            $port | Select-Object Name, Description, PrinterHostAddress, Protocol | Format-List
            
            # Netzwerk-Test falls IP-Adresse vorhanden
            if ($port.PrinterHostAddress -and $port.PrinterHostAddress -ne "") {
                Write-Host "`n6. NETZWERK-KONNEKTIVITÄT" -ForegroundColor Cyan
                Write-Host "=========================" -ForegroundColor Cyan
                Write-Host "Teste Verbindung zu: $($port.PrinterHostAddress)" -ForegroundColor Yellow
                
                $pingResult = Test-Connection -ComputerName $port.PrinterHostAddress -Count 3 -Quiet -ErrorAction SilentlyContinue
                if ($pingResult) {
                    Write-Host "✓ Netzwerkverbindung erfolgreich" -ForegroundColor Green
                }
                else {
                    Write-Host "✗ Netzwerkverbindung fehlgeschlagen" -ForegroundColor Red
                }
            }
        }
        
        Write-Host "`n7. TREIBER-INFORMATIONEN" -ForegroundColor Cyan
        Write-Host "========================" -ForegroundColor Cyan
        $driver = Get-PrinterDriver | Where-Object { $_.Name -eq $printer.DriverName }
        if ($driver) {
            $driver | Select-Object Name, Manufacturer, DriverVersion, PrintProcessor | Format-List
        }
        
        # GPO-Informationen abrufen
        $gpoInfo = Get-PrinterGPOInformation -PrinterName $PrinterName
        
        # Aktionen anbieten
        Show-PrinterActions -PrinterName $PrinterName
        
    }
    catch {
        Write-Host "`nFehler bei der Diagnose: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Mögliche Ursachen:" -ForegroundColor Yellow
        Write-Host "- Drucker existiert nicht" -ForegroundColor Gray
        Write-Host "- Keine Berechtigung" -ForegroundColor Gray
        Write-Host "- Drucker ist nicht verfügbar" -ForegroundColor Gray
    }
    
    Write-Host "`nDrücken Sie eine beliebige Taste, um fortzufahren..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

function Show-PrinterActions {
    param([string]$PrinterName)
    
    Write-Host "`n=== VERFÜGBARE AKTIONEN ===" -ForegroundColor Magenta
    Write-Host "1. Drucker neu starten"
    Write-Host "2. Alle Druckaufträge löschen"
    Write-Host "3. Drucker online setzen"
    Write-Host "4. GPO aktualisieren (gpupdate)"
    Write-Host "5. GP-Resultate anzeigen"
    Write-Host "6. Zurück zur Druckerauswahl"
    Write-Host "Q. Beenden"
    
    do {
        $action = Read-Host "`nWählen Sie eine Aktion"
        
        switch ($action.ToUpper()) {
            '1' {
                try {
                    Restart-Printer -Name $PrinterName
                    Write-Host "✓ Drucker wird neu gestartet..." -ForegroundColor Green
                }
                catch {
                    Write-Host "✗ Fehler beim Neustart: $($_.Exception.Message)" -ForegroundColor Red
                }
            }
            '2' {
                try {
                    $jobs = Get-PrintJob -PrinterName $PrinterName
                    if ($jobs) {
                        $jobs | Remove-PrintJob
                        Write-Host "✓ Alle Druckaufträge wurden gelöscht" -ForegroundColor Green
                    }
                    else {
                        Write-Host "Keine Druckaufträge zum Löschen vorhanden" -ForegroundColor Yellow
                    }
                }
                catch {
                    Write-Host "✗ Fehler beim Löschen: $($_.Exception.Message)" -ForegroundColor Red
                }
            }
            '3' {
                try {
                    Set-Printer -Name $PrinterName -WorkOffline $false
                    Write-Host "✓ Drucker wurde online gesetzt" -ForegroundColor Green
                }
                catch {
                    Write-Host "✗ Fehler: $($_.Exception.Message)" -ForegroundColor Red
                }
            }
            '4' {
                Write-Host "Führe Gruppenrichtlinien-Update aus..." -ForegroundColor Yellow
                try {
                    $result = & gpupdate /force 2>&1
                    Write-Host "✓ GPUpdate ausgeführt" -ForegroundColor Green
                    Write-Host "Bitte warten Sie 30 Sekunden und prüfen Sie den Drucker erneut." -ForegroundColor Yellow
                }
                catch {
                    Write-Host "✗ Fehler beim GPUpdate: $($_.Exception.Message)" -ForegroundColor Red
                }
            }
            '5' {
                Write-Host "Grupppenrichtlinien-Resultate:" -ForegroundColor Yellow
                try {
                    $gpresult = & gpresult /r 2>&1
                    Write-Host $gpresult -ForegroundColor Gray
                }
                catch {
                    Write-Host "✗ Fehler beim Abrufen der GP-Resultate: $($_.Exception.Message)" -ForegroundColor Red
                }
            }
            '6' {
                Start-PrinterDiagnose
                return
            }
            'Q' {
                Write-Host "Programm beendet." -ForegroundColor Yellow
                return
            }
            default {
                Write-Host "Ungültige Auswahl! Bitte wählen Sie 1-6 oder Q." -ForegroundColor Red
            }
        }
    } while ($action.ToUpper() -ne 'Q' -and $action -ne '6')
}

# Hauptprogramm starten
Start-PrinterDiagnose
