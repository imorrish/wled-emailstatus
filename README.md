# wled-emailstatus
Check 365 Exchange mailbox using PowerShell with Graph API for spcific email type and set WLed status light to notify.

Add a profile to Wled for the new and clear status effects.

Reguires an Entra App Registration (App Management or GA permissions)
- Create "PnP Exchange Delegated Access" App Registration
- Grant Delegated App permissions for Mail.Read, User.Read

Run the script using pwsh (I'm running it on MacOS).
