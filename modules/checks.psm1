# Line to force strict mode and avoid false manipulation of data
Set-StrictMode -Version latest

# Check if the current user has administrative rights
function Find-UserHasAdminRights {
    [OutputType([bool])]
    param ()

    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        return $false
    }
    return $true
}


# TODO: refactor to check if user has REQUIRED module installed, use static array of string
# Check if the current has Microsoft Graph modules installed
function Find-UserHasMicrosoftGraphInstalled {
    [OutputType([bool])]
    param ()

    # Get all installed Microsoft.Graph modules
    [System.Object]$graph_modules = Get-InstalledModule -Name "Microsoft.Graph.*" -ErrorAction SilentlyContinue

    if ($null -eq $graph_modules){
        . (Join-Path $PSScriptRoot "..\scripts\install_graph_modules.ps1")
        return $false
    }
    return $true
}


function Wait-ApiResponse {
    <#
.SYNOPSIS
    # Wait a certain time to get a response from the API
.NOTES
    Name: Wait-ApiResponse
    Author: Jocelin THUMELIN
    Version: 1.0
    DateCreated: 13.01.2025
 
.PARAMETER max_sec_wait
    (Required) Time in second before the function return Null

.PARAMETER delay
    (Optional) Delay between the API call

.PARAMETER api_request
    (Required) API call, need to be between {}

.EXAMPLE
    $response = Wait-ApiResponse -max_sec_wait 30 -api_request {
	    Get-ApiResponse -user_id $my_id -parameter $my_param
    }
 
.INPUTS
    Int,
    Int,
    ScriptBlock
    
.OUTPUTS
    Object returned by the API

    or
    
    Null

.LINK
    https://www.sectioninformatique.ch
#>
[CmdletBinding()]
param (
    [int]$max_sec_wait,
    [int]$delay = 1,
    [scriptblock]$api_request
)

BEGIN {}

PROCESS {
    
    [int]$current_sec = 0
    $api_response = $null
    
    if ($delay -lt 1) {
        $delay == 1
    }

    do {
        Start-Sleep -Seconds $delay
        
        try {
            $api_response = & $api_request
        } catch {
            Write-Host "Error during API request: $_" -ForegroundColor Yellow
        }

        $current_sec += $delay
        Write-Host "Waiting API response: $($current_sec) sec"

    } while ($null -eq $($api_response) -and $current_sec -lt $max_sec_wait)
        
    if ($null -eq $($api_response)) {
        Write-Host "Timed out ($max_sec_wait sec) No response from API" -ForegroundColor Red
        return $null
    }

    return $api_response
}

END {}
}



# Visible functions from this module
Export-ModuleMember -Function Find-UserHasAdminRights
Export-ModuleMember -Function Find-UserHasMicrosoftGraphInstalled
Export-ModuleMember -Function Wait-ApiResponse

