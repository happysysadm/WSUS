#author: happysysadm.com

$action = $false

$now = Get-Date

$comments = "Today is $($now.ToLongDateString())`n"

$WSUSServerParams = @{
    Name   = 'wsusserver.contoso.com'
    Port   = 8530
    UseSSL = $false
}

# Moved these to the top as others may want to tweak as necessary
$SyncDelay = 13 # How many days after Patch Tuesday should we wait before syncing WSUS
$WSUSCleanUpDay = 7 # What numerical day of the month whould the WSUS cleanup script run?


# Changed delay settings to use objects, as that is the most flexible
$DelaySettings = @()
$DelaySettings += [pscustomobject]@{
    Name          = 'Immediate'
    # Now multiple collections can share the same delay settings without adding multiple checks
    Collections   = 'Standard', 'NonCritical'
    ApprovalDelay = 1
}


$DelaySettings += [pscustomobject]@{
    Name          = 'OneWeek'
    Collections   = 'Touchy', 'Critical'
    ApprovalDelay = 7
}

$firstOfThisMonth = (Get-Date -Day 1 )

switch ( $firstOfThisMonth.DayOfWeek ) {

    "Sunday" {$thisPatchTuesday = $firstOfThisMonth.AddDays(9)}

    "Monday" {$thisPatchTuesday = $firstOfThisMonth.AddDays(8)}

    "Tuesday" {$thisPatchTuesday = $firstOfThisMonth.AddDays(7)}

    "Wednesday" {$thisPatchTuesday = $firstOfThisMonth.AddDays(13)}

    "Thursday" {$thisPatchTuesday = $firstOfThisMonth.AddDays(12)}

    "Friday" {$thisPatchTuesday = $firstOfThisMonth.AddDays(11)}

    "Saturday" {$thisPatchTuesday = $firstOfThisMonth.AddDays(10)}

}

if ($now.date -le $thisPatchTuesday.date) {
    $patchTuesday = $thisPatchTuesday
   
}
else {
    $firstOfNextMonth = (Get-Date -Day 1 -Month ((Get-Date).AddMonths(1).Month) )

    switch ( $firstOfNextMonth.DayOfWeek ) {

        "Sunday" {$patchTuesday = $firstOfNextMonth.AddDays(9); break}

        "Monday" {$patchTuesday = $firstOfNextMonth.AddDays(8); break}

        "Tuesday" {$patchTuesday = $firstOfNextMonth.AddDays(7); break}

        "Wednesday" {$patchTuesday = $firstOfNextMonth.AddDays(13); break}

        "Thursday" {$patchTuesday = $firstOfNextMonth.AddDays(12); break}

        "Friday" {$patchTuesday = $firstOfNextMonth.AddDays(11); break}

        "Saturday" {$patchTuesday = $firstOfNextMonth.AddDays(10); break}

    }
}

$SyncDay = (Get-Date -Date $patchTuesday).AddDays($SyncDelay)



switch ($now.Date) {
    $patchTuesday.Date {
        $action = $true
        $comments += "==> It's patch Tuesday!`n"
        $comments += "Next Sync will happen in $((New-TimeSpan -Start $now.date -End $SyncDay.date).days) days on $($SyncDay.ToLongDateString())`n"
    }

    $SyncDay.Date {
        $action = $true
        $comments += "Next Patch Tuesday is in $((New-TimeSpan -Start $now.date -End $patchTuesday.date).days) days on $($patchTuesday.ToLongDateString())`n"
        $comments += "==> It's sync day! - synching WSUS with Microsoft!`n"

        $startTime = (get-date -f dd-MM-yyyy)
        (Get-WsusServer).GetSubscription().StartSynchronization()

    }

    default {
        $action = $false
        $comments += "Next Patch Tuesday is in $((New-TimeSpan -Start $now.date -End $patchTuesday.date).days) days on $($patchTuesday.ToLongDateString())`n"
        $comments += "Next Sync will happen in $((New-TimeSpan -Start $now.date -End $SyncDay.date).days) days on $($SyncDay.ToLongDateString())`n"
    }
}

# Getting this once now, rather than for each iteration....
$wsus = Get-WsusServer @WSUSServerParams
$allupdates = $wsus.GetUpdates() 
$alltargetgroups = $wsus.GetComputerTargetGroups()
$NeededUpdates = Get-WsusUpdate -Approval Unapproved -Status FailedOrNeeded

foreach ( $Schedule in $DelaySettings ) {

    if ($now.Date -eq (Get-Date -Date $SyncDay).AddDays($Schedule.ApprovalDelay).date) {
        foreach ($Group in $Schedule.Collections) {
            $action = $true
            $comments += "==> It is $($Schedule.ApprovalDelay) days after Sync Day. Approving updates for $($Group)"

            #$NeededUpdates | Approve-WsusUpdate -Action Install -TargetGroupName $Group -Verbose
        }

        $startTime = (get-date -f dd-MM-yyyy)
    }
    else {
        $comments += "Next approval for the $($Schedule.Name) schedule will happen in $((New-TimeSpan -Start $now.date -End ((Get-Date -Date $SyncDay).AddDays($Schedule.ApprovalDelay).date)).days) days on $(((Get-Date -Date $SyncDay).AddDays($Schedule.ApprovalDelay).date).ToLongDateString())`n"
    }
}


if ($now.day -eq $WSUSCleanUpDay) {
    $action = $true
    $comments += "==> Today is WSUS monthly clean up day`n"; $action = $true

}

else {

    $comments += "Next WSUS monthly clean up will happen in $((New-TimeSpan -Start $now.date -End $(Get-Date -Day $WSUSCleanUpDay -Month $($now.Month + 1) -OutVariable datenextcleanup).Date).Days) days on $($datenextcleanup.ToLongDateString())`n"

}

$comments

if (!$action) {$comments += "<i style='color:red'>No actions to be done today</i>`n"}

$commentshtml = "<p style='color:blue'>" + $comments.replace("`n", '<br>') + "</p>"

# I don't know if the GetTotalSummary() method will pull live data every time it is called
# If it does, then we do NOT need to get all of the target computer groups again, and
#    the line below this can be commented out
$alltargetgroups = $wsus.GetComputerTargetGroups()

$patchreport = $alltargetgroups | ForEach {

    $Group = $_.Name

    $_.GetTotalSummary() | ForEach {

        [pscustomobject]@{

            TargetGroup               = $Group

            Needed                    = ($_.NotInstalledCount + $_.DownloadedCount)

            "Installed/NotApplicable" = ($_.NotApplicableCount + $_.InstalledCount)

            NoStatus                  = $_.UnknownCount

            PendingReboot             = $_.InstalledPendingRebootCount

        }

    }

}

$params = @{
    
    'encoding'   = [System.Text.Encoding]::UTF8
		
    'To'         = 'recipient@domain.com'

    'From'       = 'sender@domain.com'

    'SmtpServer' = "smtphost"

    'BodyAsHtml' = $true

    'Subject'    = "WSUS - Patch Report"
        
    'Body'       = (($commentshtml) + "<br>" + ($patchreport | ConvertTo-Html | Out-String))
   
}

#Send-MailMessage @params
