# Line to force strict mode and avoid false manipulation of data
Set-StrictMode -Version latest
function Add-ExcelTabToChannel {
    param (
        [Parameter(Mandatory=$true)]
        [string]$TeamId,
        [Parameter(Mandatory=$true)]
        [string]$ChannelId,
        [Parameter(Mandatory=$true)]
        [string]$ExcelFilePath,
        [Parameter(Mandatory=$true)]
        [string]$TabName
    )

    # Get SharePoint site
    $mg_group_site = Wait-ApiResponse -max_sec_wait 30 -api_request {
        Get-MgGroupSite -GroupId $TeamId -SiteId "root"
    }
    
    if (-not $mg_group_site) {
        throw "Failed to get SharePoint site for team"
    }

    # Get document library
    $mg_site_drive = Wait-ApiResponse -max_sec_wait 30 -api_request {
        Get-MgSiteDrive -SiteId $mg_group_site.Id
    }

    if (-not $mg_site_drive) {
        throw "Failed to get document library"
    }

    # Validate file exists
    if (-not (Test-Path $ExcelFilePath)) {
        throw "Excel file not found at: $ExcelFilePath"
    }

    # Upload file
    try {
        $file_stream = [System.IO.File]::OpenRead($ExcelFilePath)
        $file_name = [System.IO.Path]::GetFileName($ExcelFilePath)
        
        $uploaded_file = New-MgDriveItem -DriveId $mg_site_drive.Id -Name $file_name -File @{ 
            Content = $file_stream 
        }

        # Create tab configuration
        $tab_config = @{
            displayName = $TabName
            "teamsApp@odata.bind" = "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/com.microsoft.teamspace.tab.file.staticviewer.excel"
            configuration = @{
                contentUrl = $uploaded_file.WebUrl
                websiteUrl = $uploaded_file.WebUrl
            }
        }

        # Add tab to channel
        New-MgTeamChannelTab -TeamId $TeamId -ChannelId $ChannelId -BodyParameter $tab_config

        return $true
    }
    catch {
        Write-Error "Failed to add Excel tab: $_"
        return $false
    }
    finally {
        if ($file_stream) {
            $file_stream.Close()
        }
    }
}


Export-ModuleMember -Function Add-ExcelTabToChannel



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