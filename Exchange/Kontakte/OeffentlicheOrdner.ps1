# Prüfen Sie welche Ordner tatsächlich Kontakte enthalten
Get-PublicFolder "\Adressliste" -Recurse | ForEach-Object {
    $stats = Get-PublicFolderStatistics $_.Identity
    if ($stats.ItemCount -gt 0) {
        Write-Host "Ordner: $($_.Name) - Pfad: $($_.Identity) - Items: $($stats.ItemCount)" -ForegroundColor Green
    }
}