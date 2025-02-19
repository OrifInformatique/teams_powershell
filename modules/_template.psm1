# Line to force strict mode and avoid false manipulation of data
Set-StrictMode -Version latest

function Template-Function {
    <#
.SYNOPSIS
    # This is the template for all functions
    # Required scopes => Enter here the required scopes for using this function
.NOTES
    Name: Template-Function
    Author: Jocelin THUMELIN
    Version: 1.0
    DateCreated: 10.12.2024
 
.PARAMETER example
    (Required) This is the example params

.EXAMPLE
    Template-Function -example "A example string"
 
.INPUTS
    String
    
.OUTPUTS
    Null

.LINK
    https://www.sectioninformatique.ch
#>
[CmdletBinding()]
param (
    [string]$example
)

BEGIN {}

PROCESS {

}

END {}
}


Export-ModuleMember -Function Template-Function
