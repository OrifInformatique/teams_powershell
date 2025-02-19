# This script manage the workflow of the Teams creation
# All scripts are launch in the same scope (with . dot-sourcing)

# Line to force strict mode and avoid false manipulation of data
Set-StrictMode -Version latest

# Alias
$dict_t = [System.Collections.Specialized.OrderedDictionary]

################# Paths to scripts and custom modules #################
$path_module_checks = Join-Path $PSScriptRoot "..\modules\checks.psm1"
$path_module_menu = Join-Path $PSScriptRoot "..\modules\menu.psm1"
$path_script_install_graph_modules = Join-Path $PSScriptRoot "..\scripts\install_graph_modules.ps1"
$path_script_uninstall_graph_modules = Join-Path $PSScriptRoot "..\scripts\uninstall_graph_modules.ps1"

# $path_script_connect_graph = Join-Path $PSScriptRoot "..\scripts\connect_graph.ps1"
$path_script_create_team = Join-Path $PSScriptRoot "..\scripts\create_team.ps1"

# Custom Modules
Import-Module $path_module_checks
Import-Module $path_module_menu

 # Stop the script if the user don't have admin right
if (-not (Find-UserHasAdminRights)) {
    Clear-Host
    Write-Host " You do not have administrative rights. Please run this script as an Administrator." -ForegroundColor Red
    Write-Host " Press any key to exit script..."
    [console]::ReadKey($true).Key
    exit
}

# Add script description and script path here:
$scripts_menu = New-Object $dict_t
$scripts_menu.Add("Créer une nouvelle Team pour une personne en Observation", $path_script_create_team)
$scripts_menu.Add("Installer tous les modules de Microsoft Graph", $path_script_install_graph_modules)
$scripts_menu.Add("Désinstaller tous les modules de Microsoft Graph", $path_script_uninstall_graph_modules)
$scripts_menu.Add("Exit menu", 0) # 0 stop the menu loop


################# Main loop #################
[bool]$run = $true
while ($run) {

    $user_choice = New-Menu -dictionary_menu $scripts_menu
    
    # Launch the script that the user choosed
    if ($user_choice -ne 0){
        . ($user_choice)
        continue
    }

    $run = $false
}