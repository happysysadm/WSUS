#author: happysysadm.com

$action = $false

$now = Get-Date

$comments = "Today is $($now.ToLongDateString())`n"

$WSUSServerParams = @{

    Name   = 'wsusserver'

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

    # Added the DEV collection for disposable computers

    Collections   = 'Dev'

    ApprovalDelay = 1

}

$DelaySettings += [pscustomobject]@{

    Name          = 'TwoDays'

    # Now multiple collections can share the same delay settings without adding multiple checks

    Collections   = 'Standard', 'NonCritical'

    #Changed approval delay for non-critical server to two days so we have time to see if our disposable computers got broken by last update
    ApprovalDelay = 2

}

$DelaySettings += [pscustomobject]@{

    Name          = 'OneWeek'

    Collections   = 'Touchy', 'Critical'

    ApprovalDelay = 7

}

$firstOfThisMonth = (Get-Date $now -Day 1 )

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

    $firstOfNextMonth = (Get-Date $now -Day 1 -Month ((Get-Date $now).AddMonths(1).Month) )

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

#added this condition to fix syncday calculation

if ($now.date -le $thisPatchTuesday.date.AddDays($SyncDelay)) {

    $SyncDay = (Get-Date -Date $thisPatchTuesday).AddDays($SyncDelay)

    }

else{

    $SyncDay = (Get-Date -Date $patchTuesday).AddDays($SyncDelay)

    }

#end fix of syncday calculation

switch ($now.Date) {

    $patchTuesday.Date {

        $action = $true

        $comments += "==> It's patch Tuesday!`n"

        $comments += "Next Sync will happen in $((New-TimeSpan -Start $now.date -End $SyncDay.date).days) days on $($SyncDay.ToLongDateString())`n"

    }

    $SyncDay.Date {

        $action = $true

        $comments += "Next Patch Tuesday is in $((New-TimeSpan -Start $now.date -End $thispatchTuesday.date).days) days on $($thispatchTuesday.ToLongDateString())`n"

        $comments += "==> It's sync day! - synching WSUS with Microsoft!`n"

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

foreach ($Schedule in $DelaySettings) {

    #added handling of different date conditions

    if($now.Date -eq (Get-Date -Date $thisPatchTuesday).AddDays($SyncDelay).AddDays($schedule.ApprovalDelay).Date){

                foreach ($Group in $Schedule.Collections) {

                $action = $true

                $comments += 'Case 1 - '

                $comments += "==> It is $($Schedule.ApprovalDelay) days after Sync Day. Approving updates for $($Group)`n"

                #$NeededUpdates | Approve-WsusUpdate -Action Install -TargetGroupName $Group -Verbose

            }
    
       }

    elseif(($now.date -lt (Get-Date -Date $thisPatchTuesday).AddDays($SyncDelay).AddDays($schedule.ApprovalDelay).Date) -and ($now.date -gt $thispatchTuesday)){

        $comments += 'Case 2 - '

        $comments += "Next approval for the $($Schedule.Name) schedule will happen in $((New-TimeSpan -Start $now.date -End (((Get-Date -Date $thisPatchTuesday).AddDays($SyncDelay)).AddDays($Schedule.ApprovalDelay).date)).days) days on $((((Get-Date -Date $thisPatchTuesday).AddDays($SyncDelay)).AddDays($Schedule.ApprovalDelay).date).ToLongDateString())`n"

      }

    elseif(($now.date -gt $SyncDay.AddDays($schedule.ApprovalDelay).Date)){

        $comments += 'Case 3 - '

        $comments += "Next approval for the $($Schedule.Name) schedule will happen in $((New-TimeSpan -Start $now.date -End (((Get-Date -Date $thisPatchTuesday).AddDays($SyncDelay)).AddDays($Schedule.ApprovalDelay).date)).days) days on $((((Get-Date -Date $thisPatchTuesday).AddDays($SyncDelay)).AddDays($Schedule.ApprovalDelay).date).ToLongDateString())`n"

        }

    else{

        $comments += 'Case 4 - '

        $comments += "Next approval for the $($Schedule.Name) schedule will happen in $((New-TimeSpan -Start $now.date -End (((Get-Date -Date $PatchTuesday).AddDays($SyncDelay)).AddDays($Schedule.ApprovalDelay).date)).days) days on $((((Get-Date -Date $PatchTuesday).AddDays($SyncDelay)).AddDays($Schedule.ApprovalDelay).date).ToLongDateString())`n"

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

       'encoding'=[System.Text.Encoding]::UTF8

       'To' = 'admin@yourdomain.com'

        'From' = 'wsus@yourdomain.com'

        'SmtpServer' = "smtpserver"

        'BodyAsHtml' = $true

        'Subject' = "WSUS Report"

        'Body' = (($commentshtml) + "<br>" + ($patchreport | ConvertTo-Html | Out-String))
   
        }

Send-MailMessage @params
