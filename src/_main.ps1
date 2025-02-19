# Import-Module -DisableNameChecking

# Clear terminal
Clear-Host

# Clear variable cache
Remove-Variable * -ErrorAction SilentlyContinue; Remove-Module *; $error.Clear();

# Line to force strict mode and avoid false manipulation of data
Set-StrictMode -Version latest

# # Update PowerShellGet
# Install-Module PowerShellGet -Force -AllowClobber
# # Update NuGet provider
# Install-PackageProvider NuGet -Force -Scope CurrentUser

# Entry point, launch the script manager in the same scope
. (Join-Path $PSScriptRoot ".\script_manager.ps1")

# End of the script (Close the window after any user entry)
Write-Host " Exiting Script..." -ForegroundColor Yellow
[console]::ReadKey($true).Key

