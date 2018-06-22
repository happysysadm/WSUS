<#
.SYNOPSIS
   Invoke-WSUS is a function used to manage WSUS.

.DESCRIPTION
   Invoke-WSUS is a function that is used to determine the next Patch Tuesday, to sync a WSUS server with Microsoft, to approve patches based on target WSUS groups, and to show WSUS target groups configuration.

.PARAMETER ShowApprovalGroups
   Shows the approval groups sorted in the JSON file retrieved with the SettingFile parameter.

.PARAMETER SettingFile
   Specifies the path to a valid JSON file containing the name of schedule, the corresponding WSUS target groups and
   the delay in days between synchronization and approval.

.PARAMETER ShowWSUSTargetGroups
   Show all the existing WSUS target groups and the number of computers per group.

.PARAMETER WSUSName
   The name of the WSUS server to connect to.

.PARAMETER WSUSPort
   The port of the WSUS server to use for the connection. Default is 8530.

.PARAMETER WSUSSsl
   Specifies that the WSUS server should use Secure Sockets Layer (SSL) via HTTPS to communicate with an upstream server.

.PARAMETER SyncWSUSNow
   Forces the synchronization of the WSUS server specified with the WSUSName parameter.

.PARAMETER ShowLastWSUSSync
   Shows the date of the last WSUS synchronization with Microsoft.

.PARAMETER ShowNextPatchTuesday
   Shows the date of the next Patch Tuesday aka Update Tuesday.

.PARAMETER ShowNextSyncDay
   Shows the date of the next sync based on the value specified with the SyncDelay parameter.

.PARAMETER SyncDelay
   Specifies the number of days to wait between Patch Tuesday and the day the WSUS server is synchronized.

.PARAMETER ShowNextCleanupDay
   Shows the date of the next WSUS cleanup performed used AdamJ's script.

.PARAMETER CleanupDay
   Specifies the day of the month at which the WSUS cleanup with AdamJ's script is executed.

.PARAMETER ShowAll
   Shows the approval groups sorted from the JSON setting file, the existing WSUS target groups and the dates of the
   next events (Patch Tuesday, Synchronization day, WSUS clean up day and all the future approval dates).

.PARAMETER ShowApprovalSchedules
   Shows all the future approval dates based on the settings specified in the SettingFile parameter.

.PARAMETER RunApprovalSchedules
   Runs the approvals based on the settings specified in the SettingFile parameter and also synchronizezs WSUS if it is syncday

.PARAMETER SendMail
   Send a mail wich contains the approval groups sorted from the JSON setting file, the existing WSUS target groups and the dates of the next events (Patch Tuesday, Synchronization day, WSUS clean up day and all the futureapproval dates).

.PARAMETER SMTPServer
   Specifies the SMTP server to use to send the informative email.

.PARAMETER From
   Specifies the sender address.

.PARAMETER To
   Specifies the recipient address.

.EXAMPLE
   Invoke-WSUS -ShowApprovalGroups -SettingFile 'approvaldelaysettings.json'

.EXAMPLE
   Invoke-WSUS -ShowWSUSTargetGroups -WSUSName 'WSUSserver' -WSUSPort 8530 -WSUSSSL:$false

.EXAMPLE
   Invoke-WSUS -SyncWSUSNow -WSUSName 'WSUSserver' -WSUSPort 8530 -WSUSSSL:$false

.EXAMPLE
   Invoke-WSUS -ShowLastWSUSSync -WSUSName 'WSUSserver' -WSUSPort 8530 -WSUSSSL:$false

.EXAMPLE
   Invoke-WSUS -ShowNextPatchTuesday

.EXAMPLE
   Invoke-WSUS -ShowNextSyncDay -SyncDelay 13

.EXAMPLE
   Invoke-WSUS -ShowAll -SettingFile 'approvaldelaysettings.json' -CleanupDay 7 -SyncDelay 13  -WSUSName 'WSUSserver' -WSUSPort 8530 -WSUSSSL:$false

.EXAMPLE
   Invoke-WSUS -ShowApprovalSchedules -SettingFile 'approvaldelaysettings.json' -SyncDelay 13 -WSUSName 'WSUSserver' -WSUSPort 8530

.EXAMPLE
   Invoke-WSUS -ShowApprovalSchedules -SettingFile 'approvaldelaysettings.json' -SyncDelay 13 -RunApprovalSchedules -WSUSName 'WSUSserver' -WSUSPort 8530 -WSUSSSL:$false

.NOTES
   Author:  happysysadm
   Website: http://www.happysysadm.com
   Twitter: @sysadm2010
#>

function Invoke-WSUS {

[CmdletBinding()]
Param(

    [Parameter(Mandatory=$true,ParameterSetName='Show Approval Groups')]
    [Switch]$ShowApprovalGroups,

    [Parameter(Mandatory=$true,ParameterSetName='Show Approval Groups')]
    [Parameter(Mandatory=$true,ParameterSetName='Show Approval Schedules')]
    [Parameter(Mandatory=$true,ParameterSetName='Show All')]
    [Parameter(Mandatory=$true,ParameterSetName='Send Mail')]
    [ValidateScript({ Test-Path -Path $_ -PathType Leaf})]
    [String]$SettingFile,

    [Parameter(Mandatory=$true,ParameterSetName='Show WSUS Groups')]
    [Switch]$ShowWSUSTargetGroups,

    [Parameter(Mandatory=$true,ParameterSetName='Show WSUS Groups')]
    [Parameter(Mandatory=$true,ParameterSetName='Sync WSUS')]
    [Parameter(Mandatory=$true,ParameterSetName='Show Last Sync')]
    [Parameter(Mandatory=$true,ParameterSetName='Show Approval Schedules')]
    [Parameter(Mandatory=$true,ParameterSetName='Show All')]
    [Parameter(Mandatory=$true,ParameterSetName='Send Mail')]
    [String]$WSUSName,

    [Parameter(Mandatory=$true,ParameterSetName='Show WSUS Groups')]
    [Parameter(Mandatory=$true,ParameterSetName='Sync WSUS')]
    [Parameter(Mandatory=$true,ParameterSetName='Show Last Sync')]
    [Parameter(Mandatory=$true,ParameterSetName='Show Approval Schedules')]
    [Parameter(Mandatory=$true,ParameterSetName='Show All')]
    [Parameter(Mandatory=$true,ParameterSetName='Send Mail')]
    [Int]$WSUSPort = 8530,

    [Parameter(Mandatory=$true,ParameterSetName='Show WSUS Groups')]
    [Parameter(Mandatory=$true,ParameterSetName='Sync WSUS')]
    [Parameter(Mandatory=$true,ParameterSetName='Show Last Sync')]
    [Parameter(Mandatory=$true,ParameterSetName='Show Approval Schedules')]
    [Parameter(Mandatory=$true,ParameterSetName='Show All')]
    [Parameter(Mandatory=$true,ParameterSetName='Send Mail')]
    [Switch]$WSUSSSL = $false,

    [Parameter(Mandatory=$true,ParameterSetName='Sync WSUS')]
    [Switch]$SyncWSUSNow,

    [Parameter(Mandatory=$true,ParameterSetName='Show Last Sync')]
    [Switch]$ShowLastWSUSSync,
        
    [Parameter(Mandatory=$true,ParameterSetName='Show Next Patch Tuesday')]
    [Switch]$ShowNextPatchTuesday,

    [Parameter(Mandatory=$true,ParameterSetName='Show Next Synchronization Day')]
    [Switch]$ShowNextSyncDay,

    [Parameter(Mandatory=$true,ParameterSetName='Show Next Synchronization Day')]
    [Parameter(Mandatory=$true,ParameterSetName='Show Approval Schedules')]
    [Parameter(Mandatory=$true,ParameterSetName='Show All')]
    [Parameter(Mandatory=$true,ParameterSetName='Send Mail')]
    [Int]$SyncDelay,

    [Parameter(Mandatory=$true,ParameterSetName='Show Next CleanUp Day')]
    [Switch]$ShowNextCleanupDay,

    [Parameter(Mandatory=$true,ParameterSetName='Show Next CleanUp Day')]
    [Parameter(Mandatory=$true,ParameterSetName='Show All')]
    [Parameter(Mandatory=$true,ParameterSetName='Send Mail')]
    [Int]$CleanupDay,

    [Parameter(Mandatory=$true,ParameterSetName='Show All')]
    [Switch]$ShowAll,

    [Parameter(Mandatory=$true,ParameterSetName='Show Approval Schedules')]
    [Switch]$ShowApprovalSchedules,

    [Parameter(Mandatory=$false,ParameterSetName='Show Approval Schedules')]
    [Switch]$RunApprovalSchedules,

    [Parameter(Mandatory=$true,ParameterSetName='Send Mail')]
    [Switch]$SendMail,

    [Parameter(Mandatory=$true,ParameterSetName='Send Mail')]
    [String]$SMTPServer,

    [Parameter(Mandatory=$true,ParameterSetName='Send Mail')]
    [String]$From,

    [Parameter(Mandatory=$true,ParameterSetName='Send Mail')]
    [String]$To
)

    $Now = Get-Date '25 june 2018'
    "Today is $Now"

    $MailBody = New-Object System.Collections.ArrayList

    if($ShowApprovalGroups -or $SendMail -or $ShowAll) {

        Write-Verbose -Message "Showing Approval Delay in days for each WSUS target group"

        $DelaySettings = Get-Content $SettingFile | ConvertFrom-Json
        $DelaySettings
        $MailBody.Add(($DelaySettings | ConvertTo-Html)) | Out-Null
    }

    if($ShowWSUSTargetGroups -or $SyncWSUSNow -or $ShowLastWSUSSync -or $SendMail -or $ShowAll -or $ShowApprovalSchedules) {

        $WSUSServerParams = @{

            Name   = $WSUSName
            Port   = $WSUSPort
            UseSSL = $WSUSSSL
        }

        try {

            Write-Verbose -Message "Connecting to $WSUSName"       
            $WSUS = Get-WSUSServer @WSUSServerParams
        }

        catch { 
    
            "Failed to connect to $WSUSName"
            Break 
        }
    }

    if($ShowWSUSTargetGroups -or $SendMail -or $ShowAll) {
    
        try {
    
            Write-Verbose -Message "Retrieving target groups"
            $AllTargetGroups = $WSUS.GetComputerTargetGroups()

            Write-Verbose -Message "Showing names, ids and number of computers of each target group"    
            $GroupInfo = New-Object System.Collections.ArrayList

            foreach($TargetGroup in $AllTargetGroups) {
    
                $GroupInfo.Add("$( $TargetGroup | Select Name, @{Name='Total computers'; Expression={$TargetGroup.GetComputerTargets().Count}} )") | Out-Null

                }

            $GroupInfo | Out-String -Width 160

            $MailBody.add("<table><colgroup><col/></colgroup><tr><th>*</th></tr>")
            $MailBody.Add("$($groupinfo -replace '@{','<tr><td>' -replace';','</td><td>' -replace '}','</td></tr>')")
            $MailBody.Add("</table>")

#            $MailBody.Add("$( $GroupInfo | Select {$_.Trim()} | ConvertTo-Html )") | Out-Null

        }

        catch { 
    
            Write-Error -Message "Failed to retrieve target groups from $WSUSName" 
        }

    }

    if($SyncWSUSNow) {

        try {

            Write-Verbose -Message "Synching $WSUSName with Microsoft"
            #$WSUS.GetSubscription().StartSynchronization()
        }

        catch { 
        
            Write-Error "Failed sync of $WSUSName with Microsoft" 
        }
    }

    if($ShowLastWSUSSync -or $SendMail -or $ShowAll) {

        try {

            Write-Verbose -Message "Showing timestamp of last sync of $WSUSName with Microsoft"

            $LastSync = $WSUS.GetSubscription().LastSynchronizationTime
            "Last sync was $([Int](New-TimeSpan -Start $LastSync -end (Get-Date)).TotalDays) days ago on $LastSync"
            $MailBody.Add("Last sync was $([Int](New-TimeSpan -Start $LastSync -end (Get-Date)).TotalDays) days ago on $LastSync<br>") | Out-Null
        }

        catch { 
        
            Write-Error "Failed to retrieve timestamp of last sync of $WSUSName with Microsoft" 
        }
    }

    if($ShowNextPatchTuesday -or $ShowNextSyncDay -or $ShowApprovalSchedules -or $SendMail -or $ShowAll) {

        Write-Verbose -Message "Calculating date of the next Patch Tuesday aka Update Tuesday"

        $BaseDate = ( Get-Date -Day 12 ).Date
        $PatchTuesday = $BaseDate.AddDays( 2 - [Int]$BaseDate.DayOfWeek )

        If ((Get-Date $Now) -gt $PatchTuesday) {

            $LastPatchTuesday = $PatchTuesday
            $BaseDate = $BaseDate.AddMonths( 1 )
            $PatchTuesday = $BaseDate.AddDays( 2 - [Int]$BaseDate.DayOfWeek )
        }
    }

    if($ShowNextPatchTuesday -or $SendMail -or $ShowAll) {
    
        Write-Verbose -Message "Showing date of the next Patch Tuesday aka Update Tuesday"

        $TimespanToNextPatchTuesday = [Int](New-TimeSpan -Start ((Get-Date $Now).Date) -End $PatchTuesday.Date).TotalDays

        if($TimespanToNextPatchTuesday) {
        
            $MailBody.Add("Next Patch Tuesday will be in $TimespanToNextPatchTuesday days on $($PatchTuesday.Date.ToLongDateString())<br>") | Out-Null
            "Next Patch Tuesday will be in $TimespanToNextPatchTuesday days on $($PatchTuesday.Date.ToLongDateString())"      
        }

        else {
            
            $MailBody.Add($( "Today is Patch Tuesday" | ConvertTo-Html )) | Out-Null
            "Today is Patch Tuesday"            
        }
    }

    if($ShowNextSyncDay -or $ShowApprovalSchedules -or $SendMail -or $ShowAll) {

        Write-Verbose -Message "Showing date of the next synchronization day based on date of last and next patch tuesdays"

        if ($Now.date -le $LastPatchTuesday.date.AddDays($SyncDelay)) {
 
            $SyncDay = (Get-Date -Date $LastPatchTuesday).AddDays($SyncDelay)  
        }

        else {

            $SyncDay = (Get-Date -Date $PatchTuesday).AddDays($SyncDelay)

        }

        $TimespanToNextSyncday = (New-TimeSpan -Start (Get-Date $Now).date -End $SyncDay.Date).Days
    
        if($TimespanToNextSyncday) {
    
            $MailBody.Add("Next synchronization day will be in $TimespanToNextSyncday days on $($SyncDay.Date.ToLongDateString())<br>") | Out-Null
            "Next synchronization day will be in $TimespanToNextSyncday days on $($SyncDay.Date.ToLongDateString())"      
        }

        else {
        
            $MailBody.Add("Today is Sync Day<br>") | Out-Null
            "Today is Sync Day"      

            if($RunApprovalSchedules){ 
            
                #$WSUS.GetSubscription().StartSynchronization()
                
                }
        }
    }

   if($ShowNextCleanupDay -or $SendMail -or $ShowAll) {

        if($Now -eq (Get-Date $Now -Day $CleanupDay)) {
        
            $MailBody.Add("Today is WSUS clean up day<br>") | Out-Null
            "Today is WSUS clean up day" 
            Break   
        }

        if($Now -lt (Get-Date $Now -Day $CleanupDay)) {
        
            $TimespanToNextCleanupDay = (New-TimeSpan -Start ((Get-Date $Now).date) -End (Get-Date $Now -Day $CleanupDay)).Days         
            $MailBody.Add("Next WSUS clean up day is in $TimespanToNextCleanupDay days on $((Get-Date $Now -Day $CleanupDay).ToLongDateString())<br>") | Out-Null
            "Next WSUS clean up day is in $TimespanToNextCleanupDay days on $((Get-Date $Now -Day $CleanupDay).ToLongDateString())"           
        }

        if($Now -gt (Get-Date $Now -Day $CleanupDay)) {
        
            $TimespanToNextCleanupDay = (New-TimeSpan -Start ((Get-Date $Now).date) -End ((Get-Date ((Get-Date $Now).AddMonths(1)) -Day $CleanupDay)).ToLongDateString()).Days
            $MailBody.Add("Next WSUS clean up day is in $TimespanToNextCleanupDay days on $((Get-Date ((Get-Date $Now).AddMonths(1)) -Day $CleanupDay).ToLongDateString())<br>") | Out-Null
            "Next WSUS clean up day is in $TimespanToNextCleanupDay days on $((Get-Date ((Get-Date $Now).AddMonths(1)) -Day $CleanupDay).ToLongDateString())"     
        }

    }

    if($ShowApprovalSchedules -or $SendMail -or $ShowAll) {

        if($RunApprovalSchedules) {

            Write-Verbose -Message "Retrieving needed patches to approve from $WSUSName"
        
            $AllUpdates = $WSUS.GetUpdates()            
            #$NeededUpdates = Get-WSUSUpdate -Approval Unapproved -Status FailedOrNeeded
        }
        
        foreach ($Schedule in $DelaySettings) {

            if($Now.Date -eq (Get-Date -Date $LastPatchTuesday).AddDays($SyncDelay).AddDays($schedule.ApprovalDelay).Date) {

                foreach ($Group in $Schedule.Collections) {

                    Write-Verbose -Message 'Case 1'

                    $MailBody.Add("It is $($Schedule.ApprovalDelay) days after Sync Day. If -RunApprovalSchedules is specified this will approve updates for $($Group)<br>") | Out-Null
                    "It is $($Schedule.ApprovalDelay) days after Sync Day. If -RunApprovalSchedules is specified this will approve updates for $($Group)"
                 
                    if($RunApprovalSchedules) {

                        Write-Verbose -Message "Approving WSUS patches for $group"
                        $NeededUpdates | Approve-WSUSUpdate -Action Install -TargetGroupName $Group -Verbose
                    }
                }
            }

            elseif(($Now.date -lt (Get-Date -Date $LastPatchTuesday).AddDays($SyncDelay).AddDays($schedule.ApprovalDelay).Date) -and ($Now.date -gt $LastpatchTuesday)) {

                Write-Verbose -Message 'Case 2'

                $MailBody.Add("Next approval for the $($Schedule.Name) schedule will happen in $((New-TimeSpan -Start $Now.date -End (((Get-Date -Date $LastPatchTuesday).AddDays($SyncDelay)).AddDays($Schedule.ApprovalDelay).date)).days) days on $((((Get-Date -Date $LastPatchTuesday).AddDays($SyncDelay)).AddDays($Schedule.ApprovalDelay).date).ToLongDateString())<br>") | Out-Null
                "Next approval for the $($Schedule.Name) schedule will happen in $((New-TimeSpan -Start $Now.date -End (((Get-Date -Date $LastPatchTuesday).AddDays($SyncDelay)).AddDays($Schedule.ApprovalDelay).date)).days) days on $((((Get-Date -Date $LastPatchTuesday).AddDays($SyncDelay)).AddDays($Schedule.ApprovalDelay).date).ToLongDateString())"   
            }

            elseif(($Now.date -gt $SyncDay.AddDays($Schedule.ApprovalDelay).Date)) {

                Write-Verbose -Message 'Case 3'

                $MailBody.Add("Next approval for the $($Schedule.Name) schedule will happen in $((New-TimeSpan -Start $Now.date -End (((Get-Date -Date $LastPatchTuesday).AddDays($SyncDelay)).AddDays($Schedule.ApprovalDelay).date)).days) days on $((((Get-Date -Date $LastPatchTuesday).AddDays($SyncDelay)).AddDays($Schedule.ApprovalDelay).date).ToLongDateString())<br>") | Out-Null
                "Next approval for the $($Schedule.Name) schedule will happen in $((New-TimeSpan -Start $Now.date -End (((Get-Date -Date $LastPatchTuesday).AddDays($SyncDelay)).AddDays($Schedule.ApprovalDelay).date)).days) days on $((((Get-Date -Date $LastPatchTuesday).AddDays($SyncDelay)).AddDays($Schedule.ApprovalDelay).date).ToLongDateString())"
            }

            else {

                Write-Verbose -Message 'Case 4'

                $MailBody.Add("Next approval for the $($Schedule.Name) schedule will happen in $((New-TimeSpan -Start $Now.date -End (((Get-Date -Date $PatchTuesday).AddDays($SyncDelay)).AddDays($Schedule.ApprovalDelay).date)).days) days on $((((Get-Date -Date $PatchTuesday).AddDays($SyncDelay)).AddDays($Schedule.ApprovalDelay).date).ToLongDateString())<br>") | Out-Null
                "Next approval for the $($Schedule.Name) schedule will happen in $((New-TimeSpan -Start $Now.date -End (((Get-Date -Date $PatchTuesday).AddDays($SyncDelay)).AddDays($Schedule.ApprovalDelay).date)).days) days on $((((Get-Date -Date $PatchTuesday).AddDays($SyncDelay)).AddDays($Schedule.ApprovalDelay).date).ToLongDateString())"    
            }
        }
    }

    if($SendMail) {

        $MailBody
        Send-MailMessage -SmtpServer $SMTPServer -from $From -to $To -BodyAsHtml -Subject 'WSUS Actions Report' -Body ($MailBody | Out-String)
    }
}