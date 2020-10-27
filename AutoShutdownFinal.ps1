workflow AutoShutdownFinal
{
    Param
    (
		
        [parameter(Mandatory=$false)]
        [object] $WebhookData,

        [parameter(Mandatory=$false)]
        [string] $InParam_SPOVMListName,

        [parameter(Mandatory=$false)]
        [string] $InParam_SPOLogListName
		
    )
   	
   #For Debug if required
   # $VerbosePreference = "Continue"
   # $VerbosePreference = "SilentlyContinue"

   #If there is webhook data, assume runbook has been triggered via a webhook
    if ($WebhookData -ne $null) {

        # Collect properties of WebhookData
        $WebhookName    = $WebhookData.WebhookName
        $WebhookHeaders = $WebhookData.RequestHeader
        $WebhookBody    = $WebhookData.RequestBody

        # Collect individual headers. VMList information converted from JSON.
        $InputData = ConvertFrom-Json -InputObject $WebhookData.RequestBody
        $InParam_SPOId = $InputData.ItemId
        $InParam_SPOVMListName = $InputData.ListName
        $InParam_SPOLogListName = $InputData.LogListName
        
    } 
    #else non-webhook triggered, run across all items in the SharePoint Online List provided.
    else {
        #Check that a SharePoint list has been provided
        if ($InParam_SPOVMListName -eq $null) {
            throw "If not triggered by WebHook you need to specify a SharePoint Online VM List Name 'InParam_SPOVMListName'"
        }
        Write-Output "Non-Webhook triggered Runbook started."
        $InParam_SPOId = -1
    }
	
    #Retrive the RunAs Connection details
    $Conn = Get-AutomationConnection -Name 'AzureRunAsConnection'

    #Login with the RunAs Account
    Add-AzureRMAccount -ServicePrincipal -Tenant $Conn.TenantID `
        -ApplicationId $Conn.ApplicationID -CertificateThumbprint $Conn.CertificateThumbprint 

    #Retrive the SharePointSDK Connection details
    $SPOConn = Get-AutomationConnection -Name 'SPOnline'

    $AutomationAccountName = Get-AutomationVariable -Name "CurrentAutomationAccountName"
    
    $ResourceGroupName = Get-AutomationVariable -Name "CurrentResourceGroupName"
    
    $LogEntriesKept = Get-AutomationVariable -Name "LogEntriesKept"

    #Run all code as Inline script 
    #Note: $Using prefaced for all global variables
    InlineScript {
        #CheckShutdownScheduleEntry Function.  
        #Returns True if the time now, is in the specified TimeZone, falls inside the TimeRange provide
        #Returns False if the time now, is NOT in the specified TimeZone, falls inside the TimeRange provide
        function CheckShutdownSchedule ([string]$InParamTimeRange, [string]$InParamTimeZone)
	    {
	        # Initialising variables and set up from passed 
			$rangeStart, $rangeEnd, $parsedDay = $null
	        $currentTime = (Get-Date).ToUniversalTime()
			$strCurrentTimeZone = $InParamTimeZone
			
			$TZ = [System.TimeZoneInfo]::FindSystemTimeZoneById($strCurrentTimeZone)
		    $currentTime = [System.TimeZoneInfo]::ConvertTimeFromUtc($currentTime, $TZ)
            $midnight = $currentTime.AddDays(1).Date	        

	        try {
                if($InParamTimeRange.ToLower() -eq "always") {
                    $parsedDay = Get-Date "00:00"
                    $rangeStart = $parsedDay # Defaults to midnight
	                $rangeEnd = $parsedDay.AddHours(23).AddMinutes(59).AddSeconds(59) # End of the same day
                } else {
	                # Parse ranges if they contain '~'
	                if($InParamTimeRange -like "*~*") {
	                    $timeRangeComponents = $InParamTimeRange -split "~" | foreach {$_.Trim()}
	                    if($timeRangeComponents.Count -eq 2) {
	                        $rangeStart = Get-Date $timeRangeComponents[0]
	                        $rangeEnd = Get-Date $timeRangeComponents[1]
	
	                        # Check for ranges crossing midnight
	                        if($rangeStart -gt $rangeEnd) {
                                # If current time is between the start of range and midnight tonight, interpret start time as earlier today and end time as tomorrow
                                if($currentTime -ge $rangeStart -and $currentTime -lt $midnight) {
                                    $rangeEnd = $rangeEnd.AddDays(1)
                                }
                                # Otherwise interpret start time as yesterday and end time as today   
                                else {
                                    $rangeStart = $rangeStart.AddDays(-1)
                                }
	                        }
	                    }
                        # Otherwise assume there is only record and don't need to split
	                    else {
	                        Write-Error "`tWARNING: Invalid time range format. Expects valid .Net DateTime-formatted start time and end time separated by '~'" 
	                    }
	                }
	                # Otherwise attempt to parse as a full day entry, e.g. 'Monday' or 'December 25' 
	                else {
	                    # If specified as day of week, check if today
	                    if([System.DayOfWeek].GetEnumValues() -contains $InParamTimeRange) {
	                        if($InParamTimeRange -eq (Get-Date).DayOfWeek) {
	                            $parsedDay = Get-Date "00:00"
	                        }
                            # Otherwise skip as it isn't today
	                        else {
	                            # Skip detected day of week that isn't today
	                        }
	                    }
	                    # Otherwise attempt to parse as a date, e.g. 'December 25'
	                    else {
	                        $parsedDay = Get-Date $InParamTimeRange
	                    }
	    
	                    if($parsedDay -ne $null) {
	                        $rangeStart = $parsedDay # Defaults to midnight
	                        $rangeEnd = $parsedDay.AddHours(23).AddMinutes(59).AddSeconds(59) # End of the same day
	                    }                    
	                }
                }
	        }
	        catch {
	            # Record any errors and return false by default
	            Write-Error "`tWARNING: Exception encountered while parsing time range. Details: $($_.Exception.Message). Check the syntax of entry, e.g. '<StartTime> -> <EndTime>', or days/dates like 'Sunday' and 'December 25'"   
	            return $false
	        }

	        # Check if current time falls within range for shutdown
	        if($currentTime -ge $rangeStart -and $currentTime -le $rangeEnd) {
	            return $true
	        }
	        else {
	            return $false
	        }
	
	    } # End function CheckShutdownSchedule
		
        function AddToLog ([string]$NewLog, [Object]$SPLogData) {
            
            if ($Using:InParam_SPOLogListName -ne $null -and $SPLogData -ne $null) {
            
                $LogArray = ($SPLogData.CurrLog -split "`n")
                $Log = ""
                #trim the log to last entries as per Asset variable 'LogEntriesKept'
                For ($i=0; $i -lt $SPLogData.LogEntriesKept; $i++)  {
                    $Log = "$($Log)`n$($LogArray[$i])"
                }
             
                $LogEntry = "$($NewLog.Trim())`n$($Log.Trim())"
            
                $SPData = @{
                    Title = $($SPLogData.VMName)
                    SubscriptionGUID = $($SPLogData.Sub)
                    ResourceGroup = $($SPLogData.ResourceGroup)
                    AutoActionLog = "$($LogEntry)"
                }
                if ($SPLogData.LogID -eq $null) {
                    $SPOResult = Add-SPListItem -SPConnection $Using:SPOConn -ListName $Using:InParam_SPOLogListName -ListFieldsValues $SPData
                } else {
                    $SPOResult = Update-SPListItem -SPConnection $Using:SPOConn -ListName $Using:InParam_SPOLogListName -ListItemId $SPOLogID -ListFieldsValues $SPData
                }
                Write-Verbose "Log Update Result: $($SPOResult)"
                $SPOLogs = Get-SPListItem -SPConnection $Using:SPOConn -ListName $Using:InParam_SPOLogListName
                $SPOLogs
            } else {
                $null
            }
        }

        # Set Action Constants
        $EmtpyAction = 0
        $TargetGood = 1
        $StopVM = 2
        $StartVM = 4
        $MultiVMs = 8

        #Set the SharePoint Online List Names from global variables
        $SPOVMListName = $Using:InParam_SPOVMListName

        # Get a list of all Subscriptions that the current account has access to
        # This will be used to filter out records that account does not have access to
        $AvailableSubs = Get-AzureRmSubscription -WarningAction SilentlyContinue

        #If there is no WebHook Data then get all SPO records
        if ($Using:InParam_SPOId -eq -1) {
            Write-Output "Getting all SharePoint records"
            $SPOVMs = Get-SPListItem -SPConnection $Using:SPOConn -ListName $SPOVMListName
            $AllCount = $SPOVMs.count

            # Remove VM records with a DateDeleted populated
            $SPOVMs = $SPOVMs | Where-Object {$_.DateRemoved -le "1 Jan 2000"}
            $ActiveCount = $SPOVMs.count
            Write-Output "Retrieved $($AllCount) SPO records, $($AllCount - $ActiveCount) are deleted VMs"

            # Remove VM records in invalid subscriptions
            $SPOVMs = $SPOVMs | Where-Object {$_.SubscriptionGUID -in $AvailableSubs.SubscriptionId}
            $ActiveCount = $SPOVMs.count
            Write-Output "Retrieved $($AllCount) SPO records, $($AllCount - $ActiveCount) are in invalid subscriptions"

        } 
        # Otherwise just get the record requested
        else {
            Write-Output "Getting Item ID: $($Using:InParam_SPOId)"
            $SPOVMs = Get-SPListItem -SPConnection $Using:SPOConn -ListName $SPOVMListName -ListItemId $Using:InParam_SPOId
            $AllCount = $SPOVMs.count

            # Check for valid subscription access
            if ($SPOVMs.SubscriptionGUID -notin $AvailableSubs.SubscriptionId) {
                $SPOVMs = $null
                throw "Access to required Subscription denied"
            }
        }
        
        # Retrieve all previous log data if required
        if ($Using:InParam_SPOLogListName -ne $null) {
            Write-Verbose "Get all Previous logs"
            $SPOLogs = Get-SPListItem -SPConnection $Using:SPOConn -ListName $Using:InParam_SPOLogListName
            #$SPOLogs = Get-SPListItem -SPConnection $SPOConn -ListName $SPOLogListName

        }
        # otherwise set SPOLogs variable to $null
        else {
            $SPOLogs = $null
        }
        
        # Prepare targetVMState array
        $targetVMState = @()
        #ForEach ($TZ in ([System.TimeZoneInfo]::GetSystemTimeZones())) { Write-Output "$($TZ.DisplayName) [$($TZ.Id)]"}
        # Loop through each SharePoint record
        ForEach ($SPOVM in $SPOVMs) {
            
            # Pull out the Timezone ID between [ ] in the TimeZone field
            $r = [regex] "\[([^\[]*)\]"
            $match = $r.match($($SPOVM.TimeZone))
            
            # If there is a successful match for a Timezone ID
            if ($match.Success) {
                $TZId = $match.Groups[1].Value
                # Try and get a valid TimeZone entry for the matched TimeZone Id
                try {
                    
                    $TZ = [System.TimeZoneInfo]::FindSystemTimeZoneById($TZId)
                } 
                # Otherwise assume UTC
                catch {
                    
                    $TZ = [System.TimeZoneInfo]::FindSystemTimeZoneById("UTC")
                }
            } else {
                
                $TZ = [System.TimeZoneInfo]::FindSystemTimeZoneById("UTC")
            }
            
            $CurrentTZDateTime = get-date ([System.TimeZoneInfo]::ConvertTimeFromUtc(((Get-Date).ToUniversalTime()), $TZ)) -Format "dd.MM.yyyy HH:mm"

            #Uncomment for debugging if required -->
            <#
                Write-Output "$($SPOVM.Title) - $($TZ.Id)"
                Write-Output "`tOverRide-Start: $($SPOVM.OverRide_x002d_Start)"
                Write-Output "`tOverRide-Stop: $($SPOVM.OverRide_x002d_Stop)"
                Write-Output "`tShutDownTimes: $($SPOVM.ShutdownTimeRange)"
                Write-Output "`tNotifications: $($SPOVM.NotificationEmail)"
            #>

            # Set OverRide flag to false and then starts record checks
            $OverRide = $false

            # If OverRide-Stop True then add to Target list to Stop and update OverRide flag
            # OverRide-Stop takes precedence over OverRide-Start if both are selected
            if ($SPOVM.OverRide_x002d_Stop -eq -1) {
                $target = @{
                    VMName = $($SPOVM.Title)
                    SubscriptionGUID = $($SPOVM.SubscriptionGUID)
                    ResourceGroup = $($SPOVM.ResourceGroup)
                    TargetState = "VM deallocated"
                    PrevLog = $($SPOVM.AutoActionLog)
                    SPOID = $($SPOVM.Id)
                    CurrentTime = $CurrentTZDateTime
                    ChildJob = $null
                    Status = "Pending"
                }
                $targetVMState += $target
                $OverRide = $true
            }

            # If OverRide-Start True and OverRide-Stop False then add to Target list to Start and update OverRide flag
            if (($SPOVM.OverRide_x002d_Start -eq -1) -and ($SPOVM.OverRide_x002d_Stop -ne -1)) {
                $target = @{
                    VMName = $($SPOVM.Title)
                    SubscriptionGUID = $($SPOVM.SubscriptionGUID)
                    ResourceGroup = $($SPOVM.ResourceGroup)
                    TargetState = "VM running"
                    PrevLog = $($SPOVM.AutoActionLog)
                    SPOID = $($SPOVM.Id)
                    CurrentTime = $CurrentTZDateTime
                    ChildJob = $null
                    Status = "Pending"
                }
                $targetVMState += $target
                $OverRide = $true
            }

            
            # If there is a shutdown timerange and OverRide flag is False
            if (($SPOVM.ShutdownTimeRange -ne $null) -and ($OverRide -eq $false)) {
                # Split the TimeRange and trim
                $timeRangeList = $SPOVM.ShutdownTimeRange -split "," | foreach {$_.Trim()}
                $ShutDownSched = $false

                #Loop through the list of TimeRanges found
                foreach ($timeRange in $timeRangeList) {
                    
                    #Call function CheckShutdownSchedule to see if current time specified falls inside specified range
                    $SchedCheck = CheckShutdownSchedule -InParamTimeZone $TZ.Id -InParamTimeRange $timeRange
                    
                    #If function returns True, flag for Shutdown
                    if ($SchedCheck) {
                        $ShutDownSched = $true
                    }
                }
                
                #Set the target State based on return schedule check
                if ($ShutDownSched) {
                    $targetState = "VM deallocated"
                } else {
                    $targetState = "VM running"
                }

                #Add to Target list with Target State ('VM deallocated' or 'VM running')"
                $target = @{
                    VMName = $($SPOVM.Title)
                    SubscriptionGUID = $($SPOVM.SubscriptionGUID)
                    ResourceGroup = $($SPOVM.ResourceGroup)
                    TargetState = $targetState
                    PrevLog = $($SPOVM.AutoActionLog)
                    SPOID = $($SPOVM.Id)
                    CurrentTime = $CurrentTZDateTime
                    ChildJob = $null
                    Status = "Pending"
                }
                $targetVMState += $target
            }
            

            # If there are no TimeRanges specified and neither OverRide flags are set then output message
            if (($SPOVM.ShutdownTimeRange -eq $null) -and ($OverRide -eq $false)) {
                Write-Output "`tNo Shutdown information provided"
            }
        }
        
        #Get a list of all subscriptions referenced in the TargetVMState list
        $Subs = $targetVMState.SubscriptionGUID | Select-Object -Unique
        
        #Filter out any Subscriptions the current account doesn't have access to
        $Subs = $Subs | Where-Object {$_ -in $AvailableSubs.SubscriptionId}
        
        #Set PrevSub variable used to detect a change in Subscription
        $PrevSub = ""

        # Loop until all VMs in the TargetVMState list are in the desired state
        do {
            #Set StateCheck to True, this is updated to False when a VM is found in a non-targeted state
            $StateCheck = $true
            Write-Verbose "Start Do Loop"
            # Loop through the Subscriptions in the TargetVMState List
            foreach ($Sub in $Subs) {
                Write-Verbose "Start Sub Loop"
                # If the required Subscription is not selected, select it and update PrevSub variable
                if ($Sub -ne $PrevSub) {
                    Write-Verbose "Sub changed"
                    Select-AzureRmSubscription -SubscriptionId $Sub | Out-Null
                    $PrevSub = $Sub
                }

                # Get a list of VMs within the targeted subscription
                # $SubTargets = $targetVMState | Where-Object {$_.SubscriptionGuid -eq $Sub}

                #Loop through each VM
                foreach ($target in $targetVMState) {
                    if ($target.SubscriptionGuid -eq $Sub) {
                        #Write-Output "VMName : $($target.VMName)"
                        #Write-Output "`tStatus : $($target.Status)"

                        #Setup Log Details if required
                        if ($SPOLogs -ne $null) {
                            Write-Verbose "Setting up log data"
                            $SPOCurrLog = $null
                            $SPOCurrLog = ($SPOLogs | Where-Object {(($_.Title -eq $($target.VMName)) -and ($_.SubscriptionGUID -eq $Sub))}).AutoActionLog
                            $SPOLogID = $null
                            $SPOLogID = ($SPOLogs | Where-Object {(($_.Title -eq $($target.VMName)) -and ($_.SubscriptionGUID -eq $Sub))}).Id 
                    
                            $SPLogData = @{
                                CurrLog = $SPOCurrLog
                                LogEntriesKept = $Using:LogEntriesKept
                                VMName = $($target.VMName)
                                Sub = $Sub
                                ResourceGroup = $($target.ResourceGroup)
                                LogID = $SPOLogID
                            }

                        } else {
                            Write-Verbose "Log Data setup skipped"
                            $SPLogData = $null
                        }
                    
                        # Activating, Queued, Resuming, Running, Starting, Suspending, Stopping
                        # Completed, Failed, Stopped, 
                        # Suspended
                    
                        if ($target.Status -eq "Processing") {
                            $ChildStatus = $null
                            $ChildStatus = Get-AzureRmAutomationJob -Id $target.ChildJob -ResourceGroupName $Using:ResourceGroupName -AutomationAccountName $Using:AutomationAccountName -ErrorAction SilentlyContinue
                            
                            if ($ChildStatus -ne $null) {
                                Write-Output "$($target.VMName) - Child Runbook Status: $($ChildStatus.Status)"
                            }

                            switch ($Childstatus.Status) {
                                {$_ -in ("New", "Activating", "Queued", "Resuming", "Running", "Starting", "Suspending", "Stopping")} {
                                    #Stuff is still working, Set StateCheck to False and skip checking the VM 
                                    $SkipVM = $true
                                    $StateCheck = $false
                                    $target.Status = "Processing"
                                    Write-Verbose "Stuff is still working, Set StateCheck to False and skip checking the VM"
                                }
                                {$_ -in ("Completed")} {
                                    #Jobs done, Check the VM
                                    $target.ChildJob = $null
                                    Write-Verbose "Job done, Check the VM"
                                    $target.Status = "Pending"
                                    $SPOLogs = AddToLog -NewLog "$($target.CurrentTime) : Complete" -SPLogData $SPLogData
                                }
                                {$_ -in ("Stopped") } {
                                    #Jobs done, Check the VM
                                    $target.ChildJob = $null
                                    Write-Verbose "Job stopped, Check the VM"
                                    $target.Status = "Pending"
                                    $SPOLogs = AddToLog -NewLog "$($target.CurrentTime) : Failed" -SPLogData $SPLogData
                                }
                                {$_ -in ("Suspended")} {
                                    #Something went wrong, force stop the child job and throw error
                                    Write-Verbose "Something went wrong, force stop the child job and check the VM again"
                                    Stop-AzureRmAutomationJob -Id $target.ChildJob -ResourceGroupName $Using:ResourceGroupName -AutomationAccountName $Using:AutomationAccountName 
                                    #Stop-AzureRmAutomationJob -Id $target.ChildJob -ResourceGroupName $ResourceGroupName -AutomationAccountName $AutomationAccountName -Verbose
                                    throw "Catastrophic error calling child runbook 'StartSopVM'`nStop now to avoid looping and excessive charges"
                                    $SPOLogs = AddToLog -NewLog "$($target.CurrentTime) : Error" -SPLogData $SPLogData
                                }
                                {$_ -in ("Failed")} {
                                    $SPOLogs = AddToLog -NewLog "$($target.CurrentTime) : Error" -SPLogData $SPLogData
                                    throw "Catastrophic error calling child runbook 'StartSopVM'`nStop now to avoid looping and excessive charges"
                                }
                            }
                        }
                    
                        $VM = Get-AzureRmVM -Name $target.VMName -ResourceGroupName $target.ResourceGroup -Status -WarningAction SilentlyContinue
                        
                        if ($target.Status -eq "Pending") {
                        
                            #Set required action
                            if ($VM.Statuses[1].DisplayStatus -eq $target.TargetState) {
                                $ReqAct = $TargetGood
                                $target.Status = "Completed"
                            } else {
                                $ReqAct = $EmtpyAction
                            
                                if (($VM.Statuses[1].DisplayStatus -eq "VM running") -and ($target.TargetState -eq "VM deallocated")) {
                                    $ReqAct = ($ReqAct -bor $StopVM)
                                }
                                if (($VM.Statuses[1].DisplayStatus -eq "VM deallocated") -and ($target.TargetState -eq "VM running")) {
                                    $ReqAct = ($ReqAct -bor $StartVM)
                                }
                                if ($Using:InParam_SPOId -eq -1) {
                                    $ReqAct = ($ReqAct -bor $MultiVMs)
                                }

                                if ($($VM.Statuses[1].DisplayStatus) -eq $null) {
                                    $CurrStatus = "Updating"
                                } else {
                                    $CurrStatus = $($VM.Statuses[1].DisplayStatus)
                                }
                            }

                            $params = @{
                                "InParam_SubscriptionID"=$Sub;
                                "InParam_VMName"=$($target.VMName);
                                "InParam_ResourceGroup"=$($target.ResourceGroup);
                                "InParam_Action"="To Be Set"
                            }
                            Write-Verbose "ReqAct : $($ReqAct)"
                            switch ($ReqAct) {
                            
                                {($_ -band $TargetGood) -eq $TargetGood} {
                                    Write-Output "$($target.VMName) - $($target.TargetState) : Machine all good nothing to do"
                                    $SPOLogs = AddToLog -NewLog "$($target.CurrentTime)`t : Checked VM" -SPLogData $SPLogData
                                    break
                                }
                                {(($_ -band $StartVM) -eq $StartVM)} {
                                    #Write-Output "Starting VM"
                                    $Params.InParam_Action = "Start"
                                }
                                {(($_ -band $StopVM) -eq $StopVM)} {
                                    #Write-Output "Stopping VM"
                                    $Params.InParam_Action = "Stop"
                                }
                                {!(($_ -band $StartVM) -eq $StartVM) -and !(($_ -band $StopVM) -eq $StopVM)} {
                                    #Required action undtermined, don't do anything
                                    Write-Output "$($target.VMName) - $($CurrStatus) : State changing, check again next loop"
                                    break
                                }
                                {(($_ -band $MultiVMs) -eq $MultiVMs)} {
                                    Write-Output "VM:$($target.VMName) `tCurrent: $($CurrStatus) `tTarget:$($target.TargetState) - $($Params.InParam_Action) (Child Job)"
                                    $ChildJob = Start-AzureRmAutomationRunbook -AutomationAccountName $Using:AutomationAccountName -ResourceGroupName $Using:ResourceGroupName -Name "StartStopVM" -Parameters $params
                                    $target.Status = "Processing"
                                    $target.ChildJob = $ChildJob.JobId
                                    $SPOLogs = AddToLog -NewLog "$($target.CurrentTime)`t : Stopping VM" -SPLogData $SPLogData
                                    break
                                }
                                {!(($_ -band $MultiVMs) -eq $MultiVMs)} {
                                    Write-Output "VM:$($target.VMName) `tCurrent: $($CurrStatus) `tTarget:$($target.TargetState) - $($Params.InParam_Action) (Direct)" 
                                    if ($Params.InParam_Action -eq "Stop") {
                                        Stop-AzureRmVM -Name $target.VMName -ResourceGroupName $target.ResourceGroup -Force 
                                        $SPOLogs = AddToLog -NewLog "$($target.CurrentTime)`t : Stopped VM" -SPLogData $SPLogData
                                    } 
                                    if ($Params.InParam_Action -eq "Start") {
                                        Start-AzureRmVM -Name $target.VMName -ResourceGroupName $target.ResourceGroup 
                                        $SPOLogs = AddToLog -NewLog "$($target.CurrentTime)`t : Started VM" -SPLogData $SPLogData
                                    }
                                    $target.Status = "Completing"
                                    break
                                }
                            }
                        }
                        
                        $VM = Get-AzureRmVM -Name $target.VMName -ResourceGroupName $target.ResourceGroup -Status -WarningAction SilentlyContinue
                        
                        if ($VM.Statuses[1].DisplayStatus -ne $target.TargetState) {
                            $StateCheck = $false
                            #Leave Target Status unchanged
                        }
                    }
                    
                }
                # End - ForEach ($target in $targetVMState)
            }
            if ($StateCheck -eq $false) {
                write-output "Still processing"
                sleep -Seconds 60
            }
        } until ($StateCheck -eq $true)
        Write-Output "Completed"
    }

}