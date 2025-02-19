<#
    Sven Kohler + Marc Porta
    12.01.2023
    .SYNOPSIS
    Create a task set in team from a JSON file
    .DESCRIPTION
    Connect to Microsoft.Graph and Microsoft.Team
    Get Configuration from files
    Create a new Plan in planner if needed
    Add buckets and tasks to the plan
#>
function Format-OdataUrl {
    <#
        .DESCRIPTION
        convert url in Odata 
        .PARAMETER url
        url
        .OUTPUTS
        url converted
    #>
    param (
        [string]$urlSource
    )
    PROCESS {
        $url = $urlSource.Clone()
        $url = $url.replace('%', '%25')
        $url = $url.replace('.', '%2E')
        $url = $url.replace(':', '%3A')
        $url = $url.replace('@', '%40')
        $url = $url.replace('#', '%25')
        return $url
    }
}



function New-OrifTasks{
    <#
        .DESCRIPTION
        Get configuration from JSON
        - User ID / Team
        - Tasks
        Create a Plan in Planner if needed
        Create Buckets and tasks

        .PARAMETER mg_user
        microsoft graph user
        .PARAMETER mg_team
        microsoft graph team (or group)
        .PARAMETER bucketsAndTasksJSON
        json containing buckets and task details
        .PARAMETER settingJSON
        json containg global settings
    #>
    param (
        $mg_user,
        $mg_team,
        [string]$bucketsAndTasksJSON,
        [string]$settingsJSON=''
    )
    PROCESS {

        #Get configurations from JSON Files
        $bucketsAndTasks = Get-Content -Raw $bucketsAndTasksJSON | ConvertFrom-Json
        $globalSettings = Get-Content -Raw $settingsJSON | ConvertFrom-Json

        #Get user ID and Team
        $userName = "$($mg_user.name)"
        
        Write-Host "userID: $($mg_user.id) - teamID: $($mg_team.id) - userName: $($userName)"

        #Get an existing Plan or Create it
        $planName = "Tâches $($userName)"

        #Get All Plans in the team to check if plan exist
        $plans = Get-MgGroupPlannerPlan -GroupId $mg_team.id
        $planId = ""
        foreach ($plan in $plans) {
            if ($plan.Title -eq $planName) {
                $planId = $plan.Id
                break
            }
        }
        if($planId -eq "")
        {
           $newPlan = New-MgPlannerPlan -Owner $mg_team.id -Title $planName
           $planId = $newPlan.Id
        }

        # Write-Host "planId: $($planID)"
        
        #CreateBuckets and Tasks
        $bucketsAndTasks.PSObject.Properties | ForEach-Object {
            $bucketName = $_.Name
            $tasks = $_.Value

            Write-Host "Creating bucket: $($bucketName)" -ForegroundColor Yellow
            $bucket = New-MgPlannerBucket -Name $bucketName -PlanID $planId
            # Write-Host "bucketId: $($bucket.Id)"

            # Process tasks in reverse order
            [array]::Reverse($tasks)
            foreach ($task in $tasks)
            {
                Write-Host "Creating task: $($task.title)"

                #Details
                $taskDetails = @{}
                $taskDetails.Add("Description", $task.description)
                $taskDetails.Add("PreviewType", $globalSettings.previewType)

                #References (attachements)
                $taskReferences = @{}
                foreach ($reference in $task.references){
                    $ref = @{}
                    $ref.Add("@odata.type", "microsoft.graph.plannerExternalReference")
                    $ref.Add("alias", $reference.alias)
                    #TODO: Check if there is a better way to encode the URL, like this => $url = "@" + [uri]::EscapeDataString($reference.url)
                    $url = Format-OdataUrl $reference.url
                    $taskReferences.Add($url, $ref)
                }
                $taskDetails.Add("references", $taskReferences)

                #Checklist
                $taskCheckList = @{}
               
                foreach ($check in $task.checklist) {
                    Write-Host "Add checklist:  $($check)" -ForegroundColor DarkCyan
                    $chk = @{}
                    $chk.Add("@odata.type", "microsoft.graph.plannerChecklistItem")
                    $chk.Add('title', $check)
                    $chk.Add('isChecked', $false)
                    $checkId = New-Guid
                    $taskCheckList.Add($checkId.ToString(), $chk)
                }

                $taskDetails.Add("checklist", $taskCheckList)

                #Assignments (task attributions)
                $taskAssignment = @{}
                $taskAssignment.Add("@odata.type", "microsoft.graph.plannerAssignment")
                $taskAssignment.Add("orderHint", " !")
                $taskAssignments = @{}
                $taskAssignments.Add($mg_user.Id, $taskAssignment)
        
                #Create the task
                $newTask = New-MgPlannerTask -Title $task.title -PlanId $planId -BucketId $bucket.Id -Details $taskDetails -Assignments $taskAssignments -OrderHint " !"
            }
        }
    }
}

Export-ModuleMember -Function Format-OdataUrl
Export-ModuleMember -Function New-OrifTasks
