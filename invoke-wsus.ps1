<#
.SYNOPSIS
   Invoke-Wsus is a function used to manage WSUS.
.DESCRIPTION
   Invoke-Wsus is a function that is used to determine the next Patch Tuesday, to sync a WSUS server with Microsoft,
   to approve patches based on target WSUS groups, and to show WSUS target groups configuration.
.PARAMETER ShowApprovalGroups
   Shows the approval groups sorted in the JSON file retrieved with the SettingFile parameter.
.PARAMETER SettingFile
   Specifies the path to a valid JSON file containing the name of schedule, the corresponding WSUS target groups and
   the delay in days between synchronization and approval.
.PARAMETER ShowWsusTargetGroups
   Show all the existing WSUS target groups and the number of computers per group.
.PARAMETER WsusName
   The name of the WSUS server to connect to.
.PARAMETER WsusPort
   The port of the WSUS server to use for the connection. Default is 8530.
.PARAMETER WsusSsl
   Specifies that the WSUS server should use Secure Sockets Layer (SSL) via HTTPS to communicate with an upstream server.
.PARAMETER SyncWsusNow
   Forces the synchronization of the WSUS server specified with the WsusName parameter.
.PARAMETER ShowLastWsusSync
   Shows the date of the last WSUS synchronization with Microsoft.
.PARAMETER ShowNextPatchTuesday
   Shows the date of the next Patch Tuesday aka Update Tuesday.
.PARAMETER ShowNextSyncTuesday
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
   Runs the approvals based on the settings specified in the SettingFile parameter.
.PARAMETER SendMail
   Send a mail wich contains the approval groups sorted from the JSON setting file, the existing WSUS target groups
   and the dates of the    next events (Patch Tuesday, Synchronization day, WSUS clean up day and all the future
   approval dates).
.PARAMETER SMTPServer
   Specifies the SMTP server to use to send the informative email.
.PARAMETER From
   Specifies the sender address.
.PARAMETER To
   Specifies the recipient address.
.EXAMPLE
   Invoke-Wsus -ShowApprovalGroups -SettingFile 'approvaldelaysettings.json'
.EXAMPLE
   Invoke-Wsus -ShowWsusTargetGroups -WsusName 'wsusserver' -WsusPort 8530 -WsusSSL:$false
.EXAMPLE
   Invoke-Wsus -SyncWsusNow -WsusName 'wsusserver' -WsusPort 8530 -WsusSSL:$false
.EXAMPLE
   Invoke-Wsus -ShowLastWsusSync -WsusName 'rousww0045' -WsusPort 8530 -WsusSSL:$false
.EXAMPLE
   Invoke-Wsus -ShowNextPatchTuesday
.EXAMPLE
   Invoke-Wsus -ShowNextSyncTuesday -SyncDelay 13
.EXAMPLE
   Invoke-Wsus -ShowAll -SettingFile 'approvaldelaysettings.json' -CleanupDay 7 -SyncDelay 13  -WsusName 'wsusserver' -WsusPort 8530 -WsusSSL:$false
.EXAMPLE
   Invoke-Wsus -ShowApprovalSchedules -SettingFile 'approvaldelaysettings.json' -SyncDelay 13 -WsusName 'rousww0045' -WsusPort 8530
.EXAMPLE
   Invoke-Wsus -ShowApprovalSchedules -SettingFile 'E:\Script\approvaldelaysettings.json' -SyncDelay 13 -RunApprovalSchedules -WsusName 'rousww0045' -WsusPort 8530 -WsusSSL:$false
.NOTES
   Author:  happysysadm
   Website: http://www.happysysadm.com
   Twitter: @sysadm2010
#>

function Invoke-Wsus
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true,ParameterSetName='Show Approval Groups')]
        [switch]$ShowApprovalGroups,

        [Parameter(Mandatory=$true,ParameterSetName='Show Approval Groups')]
        [Parameter(Mandatory=$true,ParameterSetName='Show Approval Schedules')]
        [Parameter(Mandatory=$true,ParameterSetName='Show All')]
        [Parameter(Mandatory=$true,ParameterSetName='Send Mail')]
        [ValidateScript({ Test-Path -Path $_ -PathType Leaf})]
        [string]$SettingFile,

        [Parameter(Mandatory=$true,ParameterSetName='Show Wsus Groups')]
        [switch]$ShowWsusTargetGroups,

        [Parameter(Mandatory=$true,ParameterSetName='Show Wsus Groups')]
        [Parameter(Mandatory=$true,ParameterSetName='Sync Wsus')]
        [Parameter(Mandatory=$true,ParameterSetName='Show Last Sync')]
        [Parameter(Mandatory=$true,ParameterSetName='Show Approval Schedules')]
        [Parameter(Mandatory=$true,ParameterSetName='Show All')]
        [Parameter(Mandatory=$true,ParameterSetName='Send Mail')]
        [string]$WsusName,

        [Parameter(Mandatory=$true,ParameterSetName='Show Wsus Groups')]
        [Parameter(Mandatory=$true,ParameterSetName='Sync Wsus')]
        [Parameter(Mandatory=$true,ParameterSetName='Show Last Sync')]
        [Parameter(Mandatory=$true,ParameterSetName='Show Approval Schedules')]
        [Parameter(Mandatory=$true,ParameterSetName='Show All')]
        [Parameter(Mandatory=$true,ParameterSetName='Send Mail')]
        [int]$WsusPort = 8530,

        [Parameter(Mandatory=$true,ParameterSetName='Show Wsus Groups')]
        [Parameter(Mandatory=$true,ParameterSetName='Sync Wsus')]
        [Parameter(Mandatory=$true,ParameterSetName='Show Last Sync')]
        [Parameter(Mandatory=$true,ParameterSetName='Show Approval Schedules')]
        [Parameter(Mandatory=$true,ParameterSetName='Show All')]
        [Parameter(Mandatory=$true,ParameterSetName='Send Mail')]
        [switch]$WsusSSL = $false,

        [Parameter(Mandatory=$true,ParameterSetName='Sync Wsus')]
        [switch]$SyncWsusNow,

        [Parameter(Mandatory=$true,ParameterSetName='Show Last Sync')]
        [switch]$ShowLastWsusSync,
        
        [Parameter(Mandatory=$true,ParameterSetName='Show Next Patch Tuesday')]
        [switch]$ShowNextPatchTuesday,

        [Parameter(Mandatory=$true,ParameterSetName='Show Next Synchronization Day')]
        [switch]$ShowNextSyncTuesday,

        [Parameter(Mandatory=$true,ParameterSetName='Show Next Synchronization Day')]
        [Parameter(Mandatory=$true,ParameterSetName='Show Approval Schedules')]
        [Parameter(Mandatory=$true,ParameterSetName='Show All')]
        [Parameter(Mandatory=$true,ParameterSetName='Send Mail')]
        [int]$SyncDelay,

        [Parameter(Mandatory=$true,ParameterSetName='Show Next CleanUp Day')]
        [switch]$ShowNextCleanupDay,

        [Parameter(Mandatory=$true,ParameterSetName='Show Next CleanUp Day')]
        [Parameter(Mandatory=$true,ParameterSetName='Show All')]
        [Parameter(Mandatory=$true,ParameterSetName='Send Mail')]
        [int]$CleanupDay,

        [Parameter(Mandatory=$true,ParameterSetName='Show All')]
        [switch]$ShowAll,

        [Parameter(Mandatory=$true,ParameterSetName='Show Approval Schedules')]
        [switch]$ShowApprovalSchedules,

        [Parameter(Mandatory=$false,ParameterSetName='Show Approval Schedules')]
        [switch]$RunApprovalSchedules,

        [Parameter(Mandatory=$true,ParameterSetName='Send Mail')]
        [switch]$SendMail,

        [Parameter(Mandatory=$true,ParameterSetName='Send Mail')]
        [string]$SMTPServer,

        [Parameter(Mandatory=$true,ParameterSetName='Send Mail')]
        [string]$From,

        [Parameter(Mandatory=$true,ParameterSetName='Send Mail')]
        [string]$To
    )

  $Now = Get-Date # '1 july 2018'

  "Today is $now"

  $MailBody = @()

  if($ShowApprovalGroups -or $SendMail -or $ShowAll){

    Write-Verbose -Message "Showing Approval Delay in days for each WSUS target group"

    $DelaySettings = Get-Content $SettingFile | ConvertFrom-Json

    $DelaySettings
    
    $MailBody += $DelaySettings | ConvertTo-Html

    }

  if($ShowWsusTargetGroups -or $SyncWsusNow -or $ShowLastWsusSync -or $SendMail -or $ShowAll){

    $WSUSServerParams = @{

        Name   = $WsusName

        Port   = $WsusPort

        UseSSL = $WsusSSL

        }

    try {

        Write-Verbose -Message "Connecting to $WsusName"
        
        $Wsus = Get-WsusServer @WSUSServerParams

        }

    catch { "Failed to connect to $WsusName"; break }

    }

  if($ShowWsusTargetGroups -or $SendMail -or $ShowAll){
    
    try{
    
        Write-Verbose -Message "Retrieving target groups"

        $AllTargetGroups = $Wsus.GetComputerTargetGroups()

        Write-Verbose -Message "Showing names, ids and number of computers of each target group"
        
        $GroupInfo = @()

        foreach($TargetGroup in $AllTargetGroups){
    
            $GroupInfo += $TargetGroup | select Name, @{Name='Total computers';Expression={$TargetGroup.GetComputerTargets().count}}

            }

        $GroupInfo | out-string -Width 160

        $MailBody += $GroupInfo | ConvertTo-Html

        }

    catch{ Write-Error -Message "Failed to retrieve target groups from $WsusName" }

    }

  if($SyncWsusNow){

    try {

        Write-Verbose -Message "Synching $WsusName with Microsoft"

        #$wsus.GetSubscription().StartSynchronization()

        }

        catch { write-error "Failed sync of $WsusName with Microsoft" }

    }

  if($ShowLastWsusSync -or $SendMail -or $ShowAll){

    try {

        Write-Verbose -Message "Showing timestamp of last sync of $WsusName with Microsoft"

        $LastSync = $Wsus.GetSubscription().LastSynchronizationTime

        "Last sync was $([int](New-TimeSpan -Start $LastSync -end (Get-Date)).TotalDays) days ago on $LastSync"

        $MailBody += "Last sync was $([int](New-TimeSpan -Start $LastSync -end (Get-Date)).TotalDays) days ago on $LastSync<br>"

        }

        catch { write-error "Failed to retrieve timestamp of last sync of $WsusName with Microsoft" }

    }

  if($ShowNextPatchTuesday -or $ShowNextSyncTuesday -or $ShowApprovalSchedules -or $SendMail -or $ShowAll){

    Write-Verbose -Message "Calculating date of the next Patch Tuesday aka Update Tuesday"

    $BaseDate = ( Get-Date -Day 12 ).Date

    $PatchTuesday = $BaseDate.AddDays( 2 - [int]$BaseDate.DayOfWeek )

    If ( (Get-Date $now) -gt $PatchTuesday )

        {

        $LastPatchTuesday = $PatchTuesday

        $BaseDate = $BaseDate.AddMonths( 1 )

        $PatchTuesday = $BaseDate.AddDays( 2 - [int]$BaseDate.DayOfWeek )

        }

    }

  if($ShowNextPatchTuesday -or $SendMail -or $ShowAll) {
    
        Write-Verbose -Message "Showing date of the next Patch Tuesday aka Update Tuesday"

        $TimespanToNextPatchTuesday = [int](New-TimeSpan -Start ((Get-Date $Now).date) -End $PatchTuesday.date).TotalDays

        if($TimespanToNextPatchTuesday){
        
            $MailBody += "Next Patch Tuesday will be in $TimespanToNextPatchTuesday days on $($PatchTuesday.date.ToLongDateString())<br>"
            
            "Next Patch Tuesday will be in $TimespanToNextPatchTuesday days on $($PatchTuesday.date.ToLongDateString())"
            
            }

        else{
            
            $MailBody += "Today is Patch Tuesday" | ConvertTo-Html

            "Today is Patch Tuesday"
            
            }
        
        }

  if($ShowNextSyncTuesday -or $ShowApprovalSchedules -or $SendMail -or $ShowAll){

    Write-Verbose -Message "Showing date of the next synchronization day based on date of last and next patch tuesdays"

    if ($Now.date -le $LastPatchTuesday.date.AddDays($SyncDelay)) {
 
        $SyncDay = (Get-Date -Date $LastPatchTuesday).AddDays($SyncDelay)
   
        }

    else{

        $SyncDay = (Get-Date -Date $PatchTuesday).AddDays($SyncDelay)

        }

    $TimespanToNextSyncday = (New-TimeSpan -Start (Get-Date $Now).date -End $SyncDay.Date).Days
    
    if($TimespanToNextSyncday){
    
        $MailBody += "Next synchronization day will be in $TimespanToNextSyncday days on $($SyncDay.date.ToLongDateString())<br>"
        
        "Next synchronization day will be in $TimespanToNextSyncday days on $($SyncDay.date.ToLongDateString())"
        
        }

        else{
        
            $MailBody += "Today is Sync Day<br>"
            
            "Today is Sync Day"
        
        }

        }

   if($ShowNextCleanupDay -or $SendMail -or $ShowAll){

        if($Now -eq (get-date $Now -Day $CleanupDay)){
        
            $MailBody += "Today is WSUS clean up day<br>"
            
            "Today is WSUS clean up day" ; break
            
            }

        if($Now -lt (get-date $Now -Day $CleanupDay)){
        
            $TimespanToNextCleanupDay = (New-TimeSpan -Start ((Get-Date $Now).date) -End (get-date $Now -Day $CleanupDay)).Days
            
            $MailBody += "Next WSUS clean up day is in $TimespanToNextCleanupDay days on $((get-date $Now -Day $CleanupDay).ToLongDateString())<br>"
            
            "Next WSUS clean up day is in $TimespanToNextCleanupDay days on $((get-date $Now -Day $CleanupDay).ToLongDateString())"
            
            }

        if($now -gt (get-date $now -Day $CleanupDay)){
        
            $TimespanToNextCleanupDay = (New-TimeSpan -Start ((Get-Date $Now).date) -End ((get-date ((get-date $Now).AddMonths(1)) -Day $CleanupDay)).ToLongDateString()).Days

            $MailBody += "Next WSUS clean up day is in $TimespanToNextCleanupDay days on $((get-date ((get-date $Now).AddMonths(1)) -Day $CleanupDay).ToLongDateString())<br>"
            
            "Next WSUS clean up day is in $TimespanToNextCleanupDay days on $((get-date ((get-date $Now).AddMonths(1)) -Day $CleanupDay).ToLongDateString())"
            
            }

   }

   if($ShowApprovalSchedules -or $SendMail -or $ShowAll){

        if($RunApprovalSchedules){

            Write-Verbose -Message "Retrieving needed patches to approve from $WsusName"
        
            $allupdates = $wsus.GetUpdates()
            
            $NeededUpdates = Get-WsusUpdate -Approval Unapproved -Status FailedOrNeeded

            }

        
        foreach ($Schedule in $DelaySettings) {

        if($Now.Date -eq (Get-Date -Date $LastPatchTuesday).AddDays($SyncDelay).AddDays($schedule.ApprovalDelay).Date){

                    foreach ($Group in $Schedule.Collections) {

                    Write-Verbose -Message 'Case 1'

                    $MailBody += "It is $($Schedule.ApprovalDelay) days after Sync Day. If -RunApprovalSchedules is specified this will approve updates for $($Group)<br>"

                    "It is $($Schedule.ApprovalDelay) days after Sync Day. If -RunApprovalSchedules is specified this will approve updates for $($Group)"

                    
                    if($RunApprovalSchedules) {

                        Write-Verbose -Message "Approving WSUS patches for $group"

                        #$NeededUpdates | Approve-WsusUpdate -Action Install -TargetGroupName $Group -Verbose

                    }

                }
    
            }

        elseif(($Now.date -lt (Get-Date -Date $LastPatchTuesday).AddDays($SyncDelay).AddDays($schedule.ApprovalDelay).Date) -and ($Now.date -gt $LastpatchTuesday)){

                Write-Verbose -Message 'Case 2'

                $MailBody += "Next approval for the $($Schedule.Name) schedule will happen in $((New-TimeSpan -Start $Now.date -End (((Get-Date -Date $LastPatchTuesday).AddDays($SyncDelay)).AddDays($Schedule.ApprovalDelay).date)).days) days on $((((Get-Date -Date $LastPatchTuesday).AddDays($SyncDelay)).AddDays($Schedule.ApprovalDelay).date).ToLongDateString())<br>"

            "Next approval for the $($Schedule.Name) schedule will happen in $((New-TimeSpan -Start $Now.date -End (((Get-Date -Date $LastPatchTuesday).AddDays($SyncDelay)).AddDays($Schedule.ApprovalDelay).date)).days) days on $((((Get-Date -Date $LastPatchTuesday).AddDays($SyncDelay)).AddDays($Schedule.ApprovalDelay).date).ToLongDateString())"
    
            }

        elseif(($Now.date -gt $SyncDay.AddDays($Schedule.ApprovalDelay).Date)){

            Write-Verbose -Message 'Case 3'

            $MailBody += "Next approval for the $($Schedule.Name) schedule will happen in $((New-TimeSpan -Start $Now.date -End (((Get-Date -Date $LastPatchTuesday).AddDays($SyncDelay)).AddDays($Schedule.ApprovalDelay).date)).days) days on $((((Get-Date -Date $LastPatchTuesday).AddDays($SyncDelay)).AddDays($Schedule.ApprovalDelay).date).ToLongDateString())<br>"

            "Next approval for the $($Schedule.Name) schedule will happen in $((New-TimeSpan -Start $Now.date -End (((Get-Date -Date $LastPatchTuesday).AddDays($SyncDelay)).AddDays($Schedule.ApprovalDelay).date)).days) days on $((((Get-Date -Date $LastPatchTuesday).AddDays($SyncDelay)).AddDays($Schedule.ApprovalDelay).date).ToLongDateString())"
    
            }

        else{

            Write-Verbose -Message 'Case 4'

            $MailBody += "Next approval for the $($Schedule.Name) schedule will happen in $((New-TimeSpan -Start $Now.date -End (((Get-Date -Date $PatchTuesday).AddDays($SyncDelay)).AddDays($Schedule.ApprovalDelay).date)).days) days on $((((Get-Date -Date $PatchTuesday).AddDays($SyncDelay)).AddDays($Schedule.ApprovalDelay).date).ToLongDateString())<br>"
            
            "Next approval for the $($Schedule.Name) schedule will happen in $((New-TimeSpan -Start $Now.date -End (((Get-Date -Date $PatchTuesday).AddDays($SyncDelay)).AddDays($Schedule.ApprovalDelay).date)).days) days on $((((Get-Date -Date $PatchTuesday).AddDays($SyncDelay)).AddDays($Schedule.ApprovalDelay).date).ToLongDateString())"
        
            }
}

    }

        if($SendMail){

        $MailBody
        
        Send-MailMessage -SmtpServer $SMTPServer -from $From -to $To -BodyAsHtml -Subject 'WSUS Actions Report' -Body ($MailBody | Out-String)

        }
}
