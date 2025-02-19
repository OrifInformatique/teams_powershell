# Line to force strict mode and avoid false manipulation of data
Set-StrictMode -Version latest


function Find-MissingRequiredModules {
        <#
    .SYNOPSIS
        # Return an array of the modules passed has parameter who aren't installed for this user
     
    .NOTES
        Name: Find-MissingRequiredModules
        Author: Jocelin THUMELIN
        Version: 1.0
        DateCreated: 10.12.2024
     
    .PARAMETER array_required_modules
        (Required) The array of module to check
     
    .EXAMPLE
        Find-MissingRequiredModules -array_required_modules array
     
    .INPUTS
        Array of String
        
    .OUTPUTS
        Array of String

    .LINK
        https://www.sectioninformatique.ch
    #>
    # [CmdletBinding([string[]])] #TODO: create an error
    param (
        [string[]]$array_required_modules
    )
        
    BEGIN {}

    PROCESS {       
        Write-Host " Checking for missing required modules"
        
        if ($null -eq $array_required_modules) {
            Write-Host " No required Array passed in params" -ForegroundColor Red
            return
        }

        [int]$array_size = $array_required_modules.Count
        [System.Collections.Generic.List[string]]$return_array = [System.Collections.Generic.List[string]]::new()

        # Loop trought the passed array of module name
        for ($i = 0; $i -lt $array_size; $i++) {
            # Write-Host " Checking module: $($array_required_modules[$i])" 

            [System.Object]$installed_module = Get-InstalledModule -Name "$($array_required_modules[$i])" -ErrorAction SilentlyContinue
            
            # If the module isn't found, add it to the return array 
            if ($null -eq $installed_module){
                # Cast [void] type to don't get return index
                [void]$return_array.Add($array_required_modules[$i])
                # Write-Host " Missing modules: $($array_required_modules[$i])" -ForegroundColor Yellow
            }
        }
        return $return_array.ToArray()
    }
    
    END {}
}


function Install-RequiredModules {
    <#
.SYNOPSIS
    # Install for the user the modules passed has parameters
 
.NOTES
    Name: Install-RequiredModules
    Author: Jocelin THUMELIN
    Version: 1.0
    DateCreated: 10.12.2024
 
.PARAMETER array_modules
    (Required) The array of module to install
 
.EXAMPLE
    Install-RequiredModules -array_modules array
 
.INPUTS
    Array of String

.LINK
    https://www.sectioninformatique.ch
#>
[CmdletBinding()]

param (
    [string[]]$array_modules
)
    
BEGIN {}

PROCESS {

    
    if ($null -eq $array_modules) {
        Write-Host " There is no module to install" -ForegroundColor Yellow
        return
    }

    # Loop trought the passed array of module name
    for ($i = 0; $i -lt $array_modules.Length; $i++) {
        try {
            Write-Host " The install of the module: $($array_modules[$i]) started, it will take some time..."

            Install-Module $($array_modules[$i]) -Force -ErrorAction Stop

        } catch {
            Write-Host " An error occurred during install: $_" -ForegroundColor Red
            EXIT 1
        }
        Write-Host " Modules: $($array_modules[$i]) has been installed." -ForegroundColor Green
    }
}

END {}
}


function Uninstall-RequiredModules {
    <#
.SYNOPSIS
    # Uninstall for the user the modules passed has parameters
 
.NOTES
    Name: Uninstall-RequiredModules
    Author: Jocelin THUMELIN
    Version: 1.0
    DateCreated: 10.12.2024
 
.PARAMETER array_modules
    (Required) The array of module to uninstall
 
.EXAMPLE
    Uninstall-RequiredModules -array_modules array
 
.INPUTS
    Array of String

.LINK
    https://www.sectioninformatique.ch
#>
[CmdletBinding()]

param (
    [string[]]$array_modules
)
    
BEGIN {}

PROCESS {

    if ($null -eq $array_modules) {
        Write-Host "There is no module to uninstall" -ForegroundColor Yellow
        return
    }

    # Loop trought the passed array of module name
    for ($i = 0; $i -lt $array_modules.Count; $i++) {
        try {
            Write-Host " The uninstall of the module: $($array_modules[$i]) started, it will take some time..."

            Uninstall-Module -Name $($array_modules[$i]) -AllVersions -Force

        } catch {
            Write-Host " An error occurred during uninstall: $_" -ForegroundColor Red
            EXIT 1
        }
        Write-Host " Modules: $($array_modules[$i]) has been uninstalled." -ForegroundColor Green
    }
}

END {}
}





Export-ModuleMember -Function Find-MissingRequiredModules
Export-ModuleMember -Function Install-RequiredModules
Export-ModuleMember -Function Uninstall-RequiredModules
