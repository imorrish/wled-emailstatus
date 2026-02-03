# Monitor inbox for PIM requests and set LED
# Required modules
$graphModuleName = "Microsoft.Graph"
$graphModuleScope = "CurrentUser"
$graphModuleForce = $true
Install-Module -Name $graphModuleName -Scope $graphModuleScope -Force:$graphModuleForce

# Connect to Microsoft Graph
# Use this App Reg for Delegated on permission
$appId = "[your Entra App registration ID]"
$tenantId = "[Your Entra Tenant ID]"
$graphScopes = @("Mail.Read")
Connect-MgGraph -ClientId $appId -TenantId $tenantId -Scopes $graphScopes

# Configuration
$pollIntervalSeconds = 60
$initialLookbackMinutes = 5
$userEmailAddress = "[your email address]"
$messageSubjectPattern = "PIM: Review*"
$messageReadState = $false
$messageTopLimit = 10
$wledBaseUrl = "http://[IP address of wled device]"
$wledStatePath = "/json/state"
$wledPresetOn = 1
$wledPresetOff = 2

# Derived values
$lastChecked = (Get-Date).AddMinutes(-$initialLookbackMinutes)
$wledUrl = "$wledBaseUrl$wledStatePath"

while ($true) {
    Write-Host "Checking inbox at $(Get-Date)..."

    # Get messages received after last check
    $messages = Get-MgUserMessage -UserId $userEmailAddress -Top $messageTopLimit |
        Where-Object {
            $_.Subject -like $messageSubjectPattern -and
            $_.IsRead -eq $messageReadState
        }

    if ($messages.Count -gt 0) {
        Write-Host "Triggering WLED preset for message: $($messages[0].Subject)"

        $body = @{
            "ps" = $wledPresetOn
        } | ConvertTo-Json
        Invoke-WebRequest -Uri $wledUrl -Method POST -Body $body -ContentType "application/json"
        #
    } else {
        $body = @{
            "ps" = $wledPresetOff
        } | ConvertTo-Json
        Invoke-WebRequest -Uri $wledUrl -Method POST -Body $body -ContentType "application/json"
    }

    $lastChecked = Get-Date
    Start-Sleep -Seconds $pollIntervalSeconds
}
