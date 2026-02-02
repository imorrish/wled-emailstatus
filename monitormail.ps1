# Monitor inbox for PIM requests and selt led
# Required modules
Install-Module -Name Microsoft.Graph -Scope CurrentUser -Force

# Connect to Microsoft Graph
# Use this App Reg for Delegated on permission
$AppID = "[your Entra App registration ID]"
$TenantID = "[Your Entra Tenant ID]"
Connect-MgGraph -ClientId "$AppID" -TenantId "$TenantID" -Scopes "Mail.Read"




# Set polling interval (e.g., every 60 seconds)
$pollInterval = 60
$lastChecked = (Get-Date).AddMinutes(-5)

# WLED preset recall URL
$wledUrl = "http://[IP address of wled device]/json/state"

while ($true) {
    Write-Host "Checking inbox at $(Get-Date)..."

    # Get messages received after last check
    $messages = Get-MgUserMessage -UserId "[your email address]" -Top 10 |
        Where-Object {
            $_.Subject -like "PIM: Review*" -and
            $_.IsRead -eq $false
        }

    if ($messages.Count -gt 0) {
        Write-Host "Triggering WLED preset for message: $($messages[0].Subject)"

        $body = @{
            "ps"=1
        } | ConvertTo-Json
        Invoke-WebRequest -Uri $wledUrl -Method POST -Body $body -ContentType "application/json"
        #
    }else{
        $body = @{
            "ps"=2
        } | ConvertTo-Json
        Invoke-WebRequest -Uri $wledUrl -Method POST -Body $body -ContentType "application/json"
    }

    $lastChecked = Get-Date
    Start-Sleep -Seconds $pollInterval
}
