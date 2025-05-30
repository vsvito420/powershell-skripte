Get-MobileDevice | ForEach-Object {
    $Device = $_;
    $Stats = Get-MobileDeviceStatistics -Identity $Device.Identity -ErrorAction SilentlyContinue; #ErrorAction SilentlyContinue verhindert Fehler, falls ein Gerät keine Statistik hat
    [PSCustomObject]@{
        User         = $Device.UserDisplayName;
        DeviceModel  = $Device.DeviceModel;
        LastSyncTime = if ($Stats) { $Stats.LastSuccessSync } else { "N/A" };
        DeviceID     = $Device.Identity
    }
} | Sort-Object LastSyncTime -Descending | Format-Table -AutoSize
