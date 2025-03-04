# Installation of required modules from Microsoft Graph. NEED admin right

# Line to force strict mode and avoid false manipulation of data
Set-StrictMode -Version latest

# Get all installed Microsoft.Graph modules
[System.Object]$graph_modules = Get-InstalledModule -Name "Microsoft.Graph.*" -ErrorAction SilentlyContinue

# If none, 
if ($null -eq $graph_modules){
    
    Write-Host "Microsoft Graph modules are not installed." -ForegroundColor Red
    
    # Ask the user if he want to install the modules (NOT case sensitive)
    [string]$user_input = ""
    while ($user_input -ne "Y" -and $user_input -ne "N") {
        $user_input = Read-Host -Prompt "Do you want to install them [Y/N]"
    }

    # Yes => Install Microsoft Graph modules (need admin right)
    if ($user_input -eq "Y") {
        try {
            Write-Host "The install of all Microsoft Graph modules started, it will take some time... (wait for the confirmation message)"

            Install-Module Microsoft.Graph -Force -ErrorAction Stop

            Write-Host "Microsoft Graph modules have been installed." -ForegroundColor Green
        } catch {
            Write-Host "An error occurred during installation: $_" -ForegroundColor Red
        }
    }
    # No => Exit the script
    else {
        Write-Host "Exiting script..." -ForegroundColor Yellow
        exit
    }
    
    # Display all modules from Microsoft Graph
    Get-InstalledModule Microsoft.Graph.*
}
else {
    Write-Host "Microsoft Graph modules are already installed." 

    # Display all modules from Microsoft Graph
    Get-InstalledModule Microsoft.Graph.*
}

Write-Host "Press any key to continue..."
[console]::ReadKey($true).Key