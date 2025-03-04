# This script uninstall all Microsoft.Graph modules found. NEED admin right

# Line to force strict mode and avoid false manipulation of data
Set-StrictMode -Version latest

# Get all installed Microsoft.Graph modules
[System.Object]$graph_modules = Get-InstalledModule -Name "Microsoft.Graph.*" -ErrorAction SilentlyContinue

# Check if any modules are found
if ($graph_modules) {
    Write-Host "The uninstall of all Microsoft Graph modules started, it will take some time... (wait for the confirmation message)"

    #Then uninstall all modules found
    foreach ($module in $graph_modules) {
        Write-Host "Uninstalling module: $($module.Name)" -ForegroundColor Yellow
        Uninstall-Module -Name $module.Name -AllVersions -Force
    }
    Write-Host "All Microsoft.Graph modules have been uninstalled." -ForegroundColor Green
} else {
    Write-Host "No Microsoft.Graph modules are installed." -ForegroundColor Red
}

Write-Host "Press any key to continue..."
[console]::ReadKey($true).Key