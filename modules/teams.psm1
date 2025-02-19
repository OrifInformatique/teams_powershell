# Line to force strict mode and avoid false manipulation of data
Set-StrictMode -Version latest

function Add-MembersToTeam {
	<#
    .SYNOPSIS
        # Add Users as Members to a Team
        # Required scope => TeamMember.ReadWriteNonOwnerRole.All | TeamMember.ReadWrite.All | Group.ReadWrite.All
     
    .NOTES
        Name: Add-MembersToTeam
        Author: Jocelin THUMELIN
        Version: 1.0
        DateCreated: 10.12.2024
     
    .PARAMETER team_id
        (Required) Team where the new user will be added
		
    .PARAMETER array_user_email
        (Required) Users email

    .EXAMPLE
        Add-MembersToTeam -team_id $team.id -user_email $user.email
     
    .INPUTS
        String
        Array of String
        
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

    foreach ($user_email in $array_user_email) {

        # Wait to avoid API throttling
        Start-Sleep -Seconds 1

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

        # While loop (limited in number of try) is there in case of API throttling 
        [bool]$member_not_added = $true
        [int]$max_try = 5
        [int]$try = 0
        do {

            # ($s is to prevent the display of Member data)
            $s = New-MgTeamMember -TeamId $team_id -BodyParameter $request_body_new_member 
            
            # Wait to avoid API throttling
            Start-Sleep -Seconds 1

            $user_found = Find-UserInTeam -user_email_or_name $user_email -team_id $team_id

            $try++

            if ($user_found) {
                Write-Host "User $($user_found.name) added to the team" -ForegroundColor Cyan
                $member_not_added = $false
            }
            if ($try -eq $max_try) {
                Write-Host "Too many try ($max_try) to add member: $($user_email)" -ForegroundColor Red
                $member_not_added = $false
            }
        } while ($member_not_added)
    }
}

END {}
}


function Edit-UsersRole {
    <#
    .SYNOPSIS
        # Change the role of users
        # Required scope => TeamMember.ReadWriteNonOwnerRole.All | TeamMember.ReadWrite.All | Group.ReadWrite.All
     
    .NOTES
        Name: Edit-UsersRole
        Author: Jocelin THUMELIN
        Version: 1.0
        DateCreated: 10.12.2024
     
    .PARAMETER team_id
        (Required) Team where the new user will be added
		
    .PARAMETER array_user_email
        (Required) Users email

    .PARAMETER target_role
        (Required) 	"owner" => max rights
					"member"=> normal rights
					"guest"	=> min rights
    .EXAMPLE
        Edit-UsersRole -team_id $team.id -user_email $user.email -target_role "owner"
     
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
        [string[]] $array_user_email,
        [string] $target_role
    )
    
BEGIN {}

PROCESS {
    $possible_roles = @("owner", "member", "guest")

    # Validate the role parameter
    if (-not $possible_roles.Contains($target_role)) {
        Write-Host "Invalid role specified: $target_role" -ForegroundColor Red
        return
    }

    # Array who will hold all Members email and ID of the passed Team
    $team_members_email = New-Object System.Collections.ArrayList
    $team_members_id = New-Object System.Collections.ArrayList

    # Get all members email and ID of the Team
    $mg_team_members = Get-MgTeamMember -TeamId $team_id -All
    foreach ($mg_member in $mg_team_members) {

        # Check if AdditionalProperties exists and contains userId
        if ($mg_member.AdditionalProperties -and $mg_member.AdditionalProperties.ContainsKey("userId")) {
            $mg_user_id = $mg_member.AdditionalProperties["userId"]

            # Get the user details
            $mg_user = Get-MgUser -UserId $mg_user_id -Property "UserPrincipalName"

            # Add Team ID and email in a list only if user is in the target list
            if ($array_user_email.Contains($mg_user.UserPrincipalName)) {
                [void]$team_members_email.Add($mg_user.UserPrincipalName)
                [void]$team_members_id.Add($mg_member.Id)
            }
        } else {
            Write-Host "No User ID found for member: $($mg_members.Id)" -ForegroundColor Red
        }
    }

    # Verify all requested users were found in the team
    foreach ($email in $array_user_email) {
        if (-not $team_members_email.Contains($email)) {
            Write-Host "User: $email isn't part of the Team" -ForegroundColor Red
            continue
        }
    }

    # Change the role of the user
    for ($i = 0; $i -lt $array_user_email.Length; $i++) {
        
        # Wait to avoid API throttling
        Start-Sleep -Seconds 1
                    
        # Check if the User has already the role wanted
        $old_role = Find-UserRoleInTeam -user_id "$($team_members_id[$i])" -team_id $team_id
        if ($old_role -eq $target_role) {
            Write-Host "$($team_members_email[$i]) has already the role: $old_role " -ForegroundColor Yellow
            continue
        }
        
        # TODO: Refactor-> Guest role can't be changed, remove and add the user back with the correct role (Use a Switch)
        # Set parameters for role change, 
        #  Switch is used because "member" is a empty array in the API request and 
        #  cause errors if there is an empty variable inside.
        $request_body = @{
            "@odata.type" = "#microsoft.graph.aadUserConversationMember"
            "user@odata.bind" = "https://graph.microsoft.com/v1.0/users('$($team_members_email[$i])')"
        }

        switch ($target_role) {

            "member" { $request_body.Add("roles", @()) }
            "owner" { $request_body.Add("roles", @("owner")) }
            "guest" { $request_body.Add("roles", @("guest")) }
            
            Default {
                Write-Host "Role: $target_role is not a possibility" -ForegroundColor Red
                return
            }
        }
        
        # While loop (limited in number of try) is there in case of API throttling 
        [bool]$role_not_changed = $true
        [int]$max_try = 5
        [int]$try = 0
        do {
            # Wait to avoid API throttling
            Start-Sleep -Seconds 1
            
            # Update the user's role in the team ($s is to prevent the display of Member data)
            $s = New-MgTeamMember -TeamId $team_id -BodyParameter $request_body

            # Wait to avoid API throttling
            Start-Sleep -Seconds 1

            # Get the role of the user
            $current_role = Find-UserRoleInTeam -user_id "$($team_members_id[$i])" -team_id $team_id

            $try++

            if ($current_role -eq $target_role -and $try -eq 3) { # 3 tries to avoid API throttling
                Write-Host "Role of User: $($team_members_email[$i]) changed: ($old_role -> $target_role)" -ForegroundColor Cyan
                $role_not_changed = $false
            }
            if ($try -eq $max_try) {
                if ($current_role -ne $target_role) { # Error
                    Write-Host "Too many try to change role of User: $($team_members_email[$i]) changed: ($current_role -> $target_role)" -ForegroundColor Red
                }
                $role_not_changed = $false
            }
        } while ($role_not_changed)
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

    Add-MembersToTeam -team_id $team_id -array_user_email $array_user_email
    Edit-UsersRole -team_id $team_id -array_user_email $array_user_email -target_role "owner"
    
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

    Add-MembersToTeam -team_id $team_id -array_user_email $array_user_email
    Edit-UsersRole -team_id $team_id -array_user_email $array_user_email -target_role "guest"
    
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
    	<#
    .SYNOPSIS
        # Try to find a user in the organization
        # Required scope => User.Read.All | Directory.Read.All
    .NOTES
        Name: Find-UserInOrganization
        Author: Jocelin THUMELIN
        Version: 1.0
        DateCreated: 10.12.2024
     
    .PARAMETER user_email_or_name
        (Required) Email or Full name of the user to find

    .EXAMPLE
        Find-UserInOrganization -user_email_or_name "newuser@email.com"
     
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
		[string]$user_email_or_name
	)
	
BEGIN {}

PROCESS {
    $mg_list_all_users = Get-MgUser -Property "Id,UserPrincipalName,DisplayName"
    
    # Check if the email exist in the organization
    foreach ($mg_user in $mg_list_all_users){
        
        # TODO: Refactor -> Find the user with only the name and display a menu if there is more than 1 result
        if ($mg_user.UserPrincipalName -eq $user_email_or_name -or $mg_user.DisplayName -eq $user_email_or_name) {

            return [PSCustomObject]@{
                id = $mg_user.Id
                email = $mg_user.UserPrincipalName
                name = $mg_user.DisplayName
            }
        }
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
