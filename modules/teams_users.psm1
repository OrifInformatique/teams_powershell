# Line to force strict mode and avoid false manipulation of data
Set-StrictMode -Version latest

function Add-MembersToTeam {
    [CmdletBinding()]
    param (
        [string] $team_id,
        [string[]] $array_user_email
    )
    
    BEGIN {}

    PROCESS {
        $added_member_email = [System.Collections.ArrayList]@()

        foreach ($user_email in $array_user_email) {
            # Initial delay to avoid API throttling
            Start-Sleep -Seconds 2

            # Check if the User is part of the organization
            $user_exist = Find-UserInOrganization -user_email_or_name $user_email 
            if ($null -eq $user_exist) {
                Write-Host "User: $user_email isn't part of the organization" -ForegroundColor Red
                continue 
            }

            # Add the User to the Team as Member
            Write-Host "Adding new Member: $user_email" -ForegroundColor Cyan
            $request_body_new_member = @{
                "@odata.type" = "#microsoft.graph.aadUserConversationMember"
                roles = @()
                "user@odata.bind" = "https://graph.microsoft.com/v1.0/users('$($user_email)')"
            }

            # Retry logic with exponential backoff
            [bool]$success = $false
            [int]$max_try = 5
            [int]$try = 0
            [int]$delay = 2

            while ($try -lt $max_try -and -not $success) {
                try {
                    # Attempt to add the member
                    $null = New-MgTeamMember -TeamId $team_id -BodyParameter $request_body_new_member -ErrorAction Stop
                    
                    # Wait with exponential backoff
                    Start-Sleep -Seconds $delay
                    
                    # Verify member was added
                    $user_found = Find-UserInTeam -user_email_or_name $user_email -team_id $team_id

                    if ($user_found) {
                        Write-Host "User: $($user_found.name) added to the team" -ForegroundColor Cyan
                        [void]$added_member_email.Add($user_email)
                        $success = $true
                    }
                }
                catch {
                    $try++
                    if ($try -lt $max_try) {
                        Write-Host "Attempt $try failed: $($_.Exception.Message)" -ForegroundColor Yellow
                        $delay *= 2  # Exponential backoff
                        Start-Sleep -Seconds $delay
                    }
                    else {
                        Write-Host "Failed to add member after $max_try attempts: $user_email" -ForegroundColor Red
                    }
                }
            }
        }
        return $added_member_email
    }

    END {}
}

function Edit-UsersRole {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string] $team_id,
        [Parameter(Mandatory=$true)]
        [string[]] $array_user_email,
        [Parameter(Mandatory=$true)]
        [ValidateSet("owner", "member", "guest")]
        [string] $target_role
    )
    
    BEGIN {
        # Validate input array
        if ($array_user_email.Count -eq 0) {
            Write-Host "No users provided to update" -ForegroundColor Yellow
            return
        }
    }

    PROCESS {
        # Create dictionaries to store member info
        $member_map = @{}
        
        # Get all members email and ID of the Team
        $mg_team_members = Get-MgTeamMember -TeamId $team_id -All
        foreach ($mg_member in $mg_team_members) {
            if (-not ($mg_member.AdditionalProperties -and 
                     $mg_member.AdditionalProperties.ContainsKey("userId"))) {
                continue
            }

            $mg_user_id = $mg_member.AdditionalProperties["userId"]
            $mg_user = Get-MgUser -UserId $mg_user_id -Property "UserPrincipalName"

            if ($array_user_email -contains $mg_user.UserPrincipalName) {
                $member_map[$mg_user.UserPrincipalName] = @{
                    Id = $mg_member.Id
                    Current_role = (Find-UserRoleInTeam -user_id $mg_member.Id -team_id $team_id)
                }
            }
        }

        # Verify all requested users were found
        $not_found_users = $array_user_email | Where-Object { -not $member_map.ContainsKey($_) }
        foreach ($email in $not_found_users) {
            Write-Host "User not found in team: $email" -ForegroundColor Red
        }

        # Process role changes
        foreach ($email in $array_user_email) {
            if (-not $member_map.ContainsKey($email)) { continue }
            
            $member_info = $member_map[$email]
            
            # Skip if already has desired role
            if ($member_info.Current_role -eq $target_role) {
                Write-Host "User $email already has role: $target_role" -ForegroundColor Yellow
                continue
            }

            # Prepare request body
            $request_body = @{
                "@odata.type" = "#microsoft.graph.aadUserConversationMember"
                "user@odata.bind" = "https://graph.microsoft.com/v1.0/users('$email')"
                "roles" = switch ($target_role) {
                    "member" { @() }
                    "owner" { @("owner") }
                    "guest" { @("guest") }
                }
            }

            # Implement retry logic with exponential backoff
            $max_attempts = 5
            $attempt = 0
            $delay = 2
            $success = $false

            while ($attempt -lt $max_attempts -and -not $success) {
                try {
                    $null = New-MgTeamMember -TeamId $team_id -BodyParameter $request_body
                    Start-Sleep -Seconds 2 # Allow time for change to propagate
                    
                    $new_role = Find-UserRoleInTeam -user_id $member_info.Id -team_id $team_id
                    if ($new_role -eq $target_role) {
                        Write-Host "Successfully changed role for $($email): $($member_info.Current_role) -> $target_role" -ForegroundColor Green
                        $success = $true
                    }
                }
                catch {
                    $attempt++
                    if ($attempt -lt $max_attempts) {
                        Write-Host "Attempt $attempt failed for $email. Retrying in $delay seconds..." -ForegroundColor Yellow
                        Start-Sleep -Seconds $delay
                        $delay *= 2
                    }
                    else {
                        Write-Host "Failed to change role for $email after $max_attempts attempts: $($_.Exception.Message)" -ForegroundColor Red
                    }
                }
            }
        }
    }

    END {}
}

function Add-OwnersToTeam {
	<#
    .SYNOPSIS
        # Add Owners to a Team
        # Required scope => TeamMember.ReadWriteNonOwnerRole.All | TeamMember.ReadWrite.All | Group.ReadWrite.All
     
    .NOTES
        Name: Add-OwnersToTeam
        Author: Jocelin THUMELIN
        Version: 1.0
        DateCreated: 10.12.2024
     
    .PARAMETER team_id
        (Required) Team where the Owners will be added
		
    .PARAMETER array_user_email
        (Required) Users email

    .EXAMPLE
        Add-OwnersToTeam -team_id $team.id -array_user_email $user.email
     
    .INPUTS
        String
        Array of String
        String
        
    .LINK
        https://www.sectioninformatique.ch
    #>
	[CmdletBinding()]
	param (
		[string] $team_id,
		[string[]] $array_user_email
	)
	
BEGIN {}

PROCESS {

    [string[]]$added_member_email = Add-MembersToTeam -team_id $team_id -array_user_email $array_user_email
    Edit-UsersRole -team_id $team_id -array_user_email $added_member_email -target_role "owner"
    
}

END {}
}


function Add-GuestsToTeam {
	<#
    .SYNOPSIS
        # Add Guests to a Team
        # Required scope => TeamMember.ReadWriteNonOwnerRole.All | TeamMember.ReadWrite.All | Group.ReadWrite.All
     
    .NOTES
        Name: Add-GuestsToTeam
        Author: Jocelin THUMELIN
        Version: 1.0
        DateCreated: 10.12.2024
     
    .PARAMETER team_id
        (Required) Team where the Guests will be added
		
    .PARAMETER array_user_email
        (Required) Users email

    .EXAMPLE
        Add-GuestsToTeam -team_id $team.id -array_user_email $user.email
     
    .INPUTS
        String
        Array of String
        String
        
    .LINK
        https://www.sectioninformatique.ch
    #>
	[CmdletBinding()]
	param (
		[string] $team_id,
		[string[]] $array_user_email
	)
	
BEGIN {}

PROCESS {

    [string[]]$added_member_email = Add-MembersToTeam -team_id $team_id -array_user_email $array_user_email
    Edit-UsersRole -team_id $team_id -array_user_email $added_member_email -target_role "guest"
    
}

END {}
}


function Find-UserInTeam {
    <#
.SYNOPSIS
    # Try to find a User in a Team
    # Required scope => User.Read.All | Directory.Read.All
.NOTES
    Name: Find-UserInTeam
    Author: Jocelin THUMELIN
    Version: 1.0
    DateCreated: 10.12.2024
 
.PARAMETER user_email_or_name
    (Required) Email or Full name of the user to find

.PARAMETER team_id
    (Required) Team where the function will try to find the User

.EXAMPLE
    Find-UserInTeam -user_email_or_name "newuser@email.com"
 
.INPUTS
    String
    
.OUTPUTS
    [PSCustomObject]@{
        id = Id
        email = UserPrincipalName
        name = DisplayName
    }

    Or

    Null

.LINK
    https://www.sectioninformatique.ch
#>
[CmdletBinding()]
param (
    [string]$user_email_or_name,
    [string]$team_id
)

BEGIN {}

PROCESS {

    if (-not $user_email_or_name -or -not $team_id) {
        Write-Host "Missing params user_email_or_name: $($user_email_or_name) or team_id: $($team_id)" -ForegroundColor Red
        return
    }

    # Get all user from the selected Team 
    # (Need to get all data, because userId is in AdditionalProperties, who is accessible only like that)
    $mg_list_team_users = Get-MgTeamMember -TeamId $team_id -All
    foreach ($mg_team_user in $mg_list_team_users){
        
        # Get the UserPrincipalName (email), because Get-MgTeamMember doesn't return it
        $mg_user_found = Get-MgUser -UserId $mg_team_user.AdditionalProperties["userId"] -Property "Id,UserPrincipalName,DisplayName"
    
        # TODO: Refactor -> Find the user with only the name and display a menu if there is more than 1 result
        # TODO: Refactor -> Same code in Find-UserInOrganization
        if ($mg_user_found.UserPrincipalName -eq $user_email_or_name -or $mg_team_user.DisplayName -eq $user_email_or_name) {
            
            return [PSCustomObject]@{
                id = $mg_user_found.Id
                email = $mg_user_found.UserPrincipalName
                name = $mg_team_user.DisplayName
            }
        }
    }
    return $null
}

END {}
}


function Find-UserInOrganization {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$user_email_or_name
    )
    
    BEGIN {
        # Handle null/empty input
        if ([string]::IsNullOrWhiteSpace($user_email_or_name)) {
            Write-Warning "Empty or null search term provided"
            return $null
        }
        # Trim and convert to lowercase for comparison
        $search_term = $user_email_or_name.Trim().ToLower()
    }

    PROCESS {
        try {
            # Use server-side filtering to improve performance
            $filter = "startsWith(userPrincipalName,'$search_term') or startsWith(displayName,'$search_term')"
            $mg_list_all_users = Get-MgUser -Filter $filter -Property "Id,UserPrincipalName,DisplayName" -ErrorAction Stop
            
            # Check for exact match first
            $exact_match = $mg_list_all_users | Where-Object { 
                ($_.UserPrincipalName -and $_.UserPrincipalName.ToLower() -eq $search_term) -or 
                ($_.DisplayName -and $_.DisplayName.ToLower() -eq $search_term) 
            } | Select-Object -First 1

            if ($exact_match) {
                return [PSCustomObject]@{
                    id = $exact_match.Id
                    email = $exact_match.UserPrincipalName
                    name = $exact_match.DisplayName
                }
            }

            # If no exact match, check for partial matches
            $partial_match = $mg_list_all_users | Where-Object { 
                ($_.DisplayName -and $_.DisplayName.ToLower().Contains($search_term))
            } | Select-Object -First 1

            if ($partial_match) {
                Write-Host "Found partial match: $($partial_match.DisplayName)" -ForegroundColor Yellow
                return [PSCustomObject]@{
                    id = $partial_match.Id
                    email = $partial_match.UserPrincipalName
                    name = $partial_match.DisplayName
                }
            }
        }
        catch {
            Write-Error "Error searching for user: $_"
        }
        
        return $null
    }

    END {}
}


function Find-UserRoleInTeam {
    <#
.SYNOPSIS
    # This function return the roles of the user in the Team: Owner | Member | Guest 
    # Required scopes => TeamMember.Read.All
.NOTES
    Name: Find-UserRoleInTeam
    Author: Jocelin THUMELIN
    Version: 1.0
    DateCreated: 06.01.2025
 
.PARAMETER user_id
    (Required) ID of the User 

.PARAMETER team_id
    (Required) ID of the Team

.EXAMPLE
    Find-UserRoleInTeam -user_id $user_id -team_id $team_id
 
.INPUTS
    String
    String
    
.OUTPUTS
    String

.LINK
    https://www.sectioninformatique.ch
#>
[CmdletBinding()]
param (
    [string]$user_id,
    [string]$team_id
    )
    
BEGIN {}
    
PROCESS {
    # TODO: Refactor-> Focus on Channel of a Team instead of Team
    # [string]$channels_id = ""
    # # Select the general channel id no parameters are given
    # if ($null -eq $channels_id) {

    # }
    # $team_members = Get-MgTeamChannelMember -TeamId $team_id -ChannelId $channel_id -All
    
    # Get all team members
    $team_members = Get-MgTeamMember -TeamId $team_id -All

    # Find a specific user by Azure AD Object ID
    $user = $team_members | Where-Object { $_.Id -eq $user_id }
    
    if ($user) {
        if ($user.Roles.Length -eq 0 -or $null -eq $user.Roles[0] -or $user.Roles[0] -eq "") {
            return "member"
        }
        return $($user.Roles)

    } else {
        return ""
    }
} 

END {}
}


Export-ModuleMember -Function Add-MembersToTeam
Export-ModuleMember -Function Add-OwnersToTeam
Export-ModuleMember -Function Add-GuestsToTeam
Export-ModuleMember -Function Edit-UsersRole
Export-ModuleMember -Function Find-UserInOrganization
Export-ModuleMember -Function Find-UserInTeam
Export-ModuleMember -Function Find-UserRoleInTeam
