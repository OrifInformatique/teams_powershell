######################### Set Up #########################
# Line to force strict mode and avoid false manipulation of data
Set-StrictMode -Version latest

# Import Custom modules
$path_module_checks = Join-Path $PSScriptRoot "..\modules\checks.psm1"
$path_module_module_manager = Join-Path $PSScriptRoot "..\modules\module_manager.psm1"
$path_module_teams_users = Join-Path $PSScriptRoot "..\modules\teams_users.psm1"
$path_module_teams_planner = Join-Path $PSScriptRoot "..\modules\teams_planner.psm1"

Import-Module $path_module_checks
Import-Module $path_module_module_manager
Import-Module $path_module_teams_users
Import-Module $path_module_teams_planner

# Required modules for this script
[string[]]$required_modules = @(
	"Microsoft.Graph.Teams",
	"Microsoft.Graph.Groups",
	"Microsoft.Graph.Users",
	"Microsoft.Graph.Planner",
	"Microsoft.Graph.Files",
	"Microsoft.Graph.Sites"
)

# Look for the missing required modules and install them if needed 
[string[]]$missing_modules = Find-MissingRequiredModules -array_required_modules $required_modules
Install-RequiredModules -array_modules $missing_modules

# TODO: Change to the minimum rights required for this script
# Scopes limit permissions required
[string]$scopes_global = "Directory.ReadWrite.All"
[string]$scopes_group = "Group.ReadWrite.All"
[string]$scopes_teams = "TeamSettings.Read.All", "TeamMember.ReadWriteNonOwnerRole.All", "TeamMember.ReadWrite.All"
[string]$scopes_planner = "Tasks.ReadWrite", "DeviceManagementApps.Read.All"
[string]$scopes_files = "Files.ReadWrite.All"
[string]$scopes_all = $scopes_global + " " + $scopes_group + " " + $scopes_teams + " " + $scopes_planner + " " + $scopes_files


######################### Entry point #########################
Clear-Host

# Connection to Microsoft Graph
Connect-MgGraph -Scopes $scopes_all -NoWelcome


######################### Ask and Find Apprentice #########################
# Ask to enter the new user email or full name 
[string]$user_input = ""
[PSCustomObject]$mg_user = $null
do{
	$user_input = Read-Host -Prompt "Enter the email or the full name of the new user"
	Write-Host "Searching for user: $user_input"
	$mg_user = Find-UserInOrganization -user_email_or_name $user_input

	if ($null -eq $mg_user) {
		Write-Host "$user_input not found" -ForegroundColor Yellow
	}
}
while ($null -eq $mg_user)
Write-Host "User: $($mg_user.name) found with email: $($mg_user.email)" -ForegroundColor Cyan

# Check if a Team with the same display name already exist
$mg_created_team = Get-MgTeam -Filter "displayName eq '$($mg_user.name)'" -Property "id,description"

if ($null -ne $($mg_created_team)) {
	Write-Host "Another Team found with the same name '$($mg_user.name)'" -ForegroundColor Red

	# Ask the user if he want to delete the Team (NOT case sensitive)
	[string]$user_input = ""
	while ($user_input -ne "Y" -and $user_input -ne "N") {
		$user_input = Read-Host -Prompt "Do you want to recreate the Team '$($mg_user.name)' [Y/N]"
	}
	# Yes => Delete the team with the same display name
	if ($user_input -eq "Y") {
		Write-Host "Delete Team '$($mg_user.name)'" -ForegroundColor Yellow
		Remove-MgGroup -GroupId $($mg_created_team.id)
	}
	# No => Exit the script
	else {
		Write-Host "Exiting script..." -ForegroundColor Yellow
		EXIT 0
	}
}


######################### Creation of the Team with Standard Channels #########################
$request_body_new_team = @{
	"template@odata.bind" = "https://graph.microsoft.com/v1.0/teamsTemplates('standard')"
	visibility = "Private"
	displayName = "$($mg_user.name)"
	description = "$($mg_user.email)"
	# Standard Channels, will be visible by any member of the Team
	channels = @(
		@{
			displayName = "Recherche de stages - emplois"
			isFavoriteByDefault = $true
			description = ""
		}
	)
	# TODO: Member settings
    memberSettings = @{
		allowCreateUpdateChannels = $false
		allowDeleteChannels = $false
		allowAddRemoveApps = $false
		allowCreateUpdateRemoveTabs = $false
		allowCreateUpdateRemoveConnectors = $false
	}
}

# Discard output to prevent the display of Team data
$null = New-MgTeam -BodyParameter $request_body_new_team

$mg_created_team = Wait-ApiResponse -max_sec_wait 30 -api_request {
    Get-MgTeam -Filter "displayName eq '$($mg_user.name)'" -Property "id,description"
}

if (-not $mg_created_team) {
    Write-Host "No response received when creating the Team" -ForegroundColor Red
	EXIT 1
}


######################### Add Private Channels to the Team #########################
$request_body_private_channel = @{
    displayName = "Formateurs"
    description = ""
    membershipType = "Private"
}

New-MgTeamChannel -TeamId $($mg_created_team.id) -BodyParameter $request_body_private_channel


# ######################### Add Tabs to Channels #########################
# # This is done like that because Teams requires secure URLs for any content it renders in tabs.

# $mg_group_site = Wait-ApiResponse -max_sec_wait 30 -api_request {
# 	# Get the SharePoint site associated with the team
# 	Get-MgGroupSite -GroupId $mg_created_team.id -SiteId "root"
# }

# Write-Host "Site ID: $($mg_group_site.Id)" -ForegroundColor Green # DEBUG
# $request_body_new_tabs = New-Object -TypeName System.Collections.ArrayList

# if ($mg_group_site) {
# 	# Get the document library (drive) ID
# 	$mg_site_drive = Wait-ApiResponse -max_sec_wait 30 -api_request {
# 		Get-MgSiteDrive -SiteId $mg_group_site.Id
# 	}

# 	if ($mg_site_drive){

# 		#### Test #### # DEBUG
# 		# [Pomy] Dufey Killian

# 		Write-Host "Site Drive ID: $($mg_site_drive.Id)" -ForegroundColor Green # DEBUG
		
# 		Write-Host "Get-MgDrive ========================================"
# 		$mg_drive = Get-MgDrive -DriveId $($mg_site_drive.Id)
# 		Write-Host "Drive ID: $($mg_drive.Id)" -ForegroundColor Green # DEBUG
# 		$mg_drive

# 		Write-Host "Get-MgDriveItem of mg_drive ========================================"
# 		# Get-MgDriveItem -DriveId $($mg_site_drive.Id) -Filter "name eq 'Documents'" -Property "id,webUrl"
# 		Get-MgDriveItem -DriveId $($mg_drive.Id) -Filter "name eq 'Documents'"
		
# 		Write-Host "Get-MgDriveItem of mg_drive_item ========================================"
# 		$mg_drive_item = Get-MgDriveItem -DriveId $($mg_site_drive.Id) -Filter "name eq 'Documents'"
# 		Write-Host "Drive Item ID: $($mg_drive_item.Id)" -ForegroundColor Green # DEBUG
# 		$mg_drive_item


# 		# [string]$file_path = "\\srv-sec\formation\Administratif\LogoEtModeles"
# 		# [string]$file_path = "C:\_Programmation_Local\PowerShell\Teams_Creation"
# 		# [string]$file_name = "_Modele_HoraireTeams.xlsx"
# 		# [string]$full_file_path = Join-Path -Path $file_path -ChildPath $file_name
		
# 		[string]$full_file_path = "C:\_Programmation_Local\PowerShell\Teams_Creation\_Modele_HoraireTeams.xlsx"
# 		[string]$file_name = [System.IO.Path]::GetFileName($full_file_path)

# 		if (-Not (Test-Path $full_file_path)) {
# 			Write-Host "File not found: $full_file_path" -ForegroundColor Red
# 			return
# 		}
# 		Write-Host "File path: $full_file_path" -ForegroundColor Magenta # DEBUG
		
# 		# Upload the file to the SharePoint document library
# 		$file_stream = [System.IO.File]::OpenRead($full_file_path)

# 		try {
# 			$uploaded_file = New-MgDriveItem -DriveId $($mg_drive_item.Id) -Name $file_name -File @{ Content = $file_stream }
# 			Write-Host "File uploaded successfully!" -ForegroundColor Green
# 			Write-Host "File Name: $($uploaded_file)"
# 			Write-Host "Web URL: $($uploaded_file)"
# 		} catch {
# 			Write-Host "Error uploading file: $_" -ForegroundColor Red
# 			Write-Host "Drive ID: $($mg_drive.Id)" -ForegroundColor Red
# 			Write-Host "File Name: $file_name" -ForegroundColor Red
# 			Write-Host "File Stream: $file_stream" -ForegroundColor Red
# 		} finally {
# 			$file_stream.Close()
# 		}
		
# 		if ($uploaded_file) {

# 			# Get the URL of the uploaded file
# 			$excel_schedule_url = $uploaded_file.WebUrl
			
# 			# TODO: Horaire (App Excel)
# 			[void]$request_body_new_tabs.Add(	
# 				@{
# 					displayName = "Horaire"
# 					"teamsApp@odata.bind" = "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/com.microsoft.teamspace.tab.file.staticviewer.excel"
# 					configuration = @{
# 						entityId = ""
# 						contentUrl = $excel_schedule_url
# 						removeUrl = ""
# 						websiteUrl = $excel_schedule_url
# 					}
# 				}
# 			)
# 		}
# 	} else {
# 		Write-Host "No response received when searching the Drive, you will need to add the the files manually" -ForegroundColor Red
# 	}
# } else {
# 	Write-Host "No response received when searching the Sharepoint, you will need to add the the files manually" -ForegroundColor Red
# }


######################### Add plans and tasks to the Planner Apps #########################
# Load the JSON file into a PowerShell object
$json_file_path = Join-Path $PSScriptRoot "..\resources\planner_obs_data.json"
$json_file_path_settings = Join-Path $PSScriptRoot "..\resources\planner_obs_settings.json"

New-OrifTasks -mg_user $mg_user -mg_team $mg_created_team -bucketsAndTasksJSON $json_file_path -settingsJSON $json_file_path_settings


######################### Add Owner and Apprentice to the Team #########################
# List of Owners email of the new team
$list_owners = @(
	"didier.viret@sectioninformatique.ch",
	"christopher.welte@sectioninformatique.ch",
	"raphael.schmutz@sectioninformatique.ch",
	"sven.kohler@sectioninformatique.ch",
	"teresa.valente@sectioninformatique.ch",
	"frederic.schmocker@sectioninformatique.ch"
)

Start-Sleep -Seconds 1
	
Add-OwnersToTeam -team_id $($mg_created_team.id) -array_user_email $list_owners
Add-MembersToTeam -team_id $($mg_created_team.id) -array_user_email $($mg_user.email)
	

# Disconnect from Microsoft Graph end of script
Disconnect-MgGraph    

Write-Host "Press any key to continue..."
[console]::ReadKey($false).Key



#### Test #### DEBUG
# [Pomy] Dufey Killian