# Connect to Microsoft Graph
Connect-Graph

# return all users
Get-MgUser

# Give the current user perission
Find-MgGraphPermission sites -PermissionType Delegated

# Disconnect to Microsoft Graph
Disconnect-Graph

Write-Host " Press any key to continue..."
[console]::ReadKey($true).Key