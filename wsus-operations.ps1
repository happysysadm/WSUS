#author: happysysadm.com

$Action = $false
$Now = Get-Date
$Comments = New-Object System.Collections.ArrayList

$Comments.Add("Today is $($Now.ToLongDateString())`n")

$WSUSServerParams = @{

    Name   = 'wsusserver'
    Port   = 8530
    UseSSL = $false
}

# Moved these to the top as others may want to tweak as necessary
$SyncDelay = 13 # How many days after Patch Tuesday should we wait before syncing WSUS
$WSUSCleanUpDay = 7 # What numerical day of the month whould the WSUS cleanup script run?

# Changed delay settings to use objects, as that is the most flexible
$DelaySettings = New-Object System.Collections.ArrayList

$DelaySettings.Add(

    [PSCustomObject]@{

        # Added the DEV collection for disposable computers
        Name          = 'Immediate'
        Collections   = 'Dev'
        ApprovalDelay = 1
    }
)

# Now multiple collections can share the same delay settings without adding multiple checks
# Changed approval delay for non-critical server to two days so we have time to see if our disposable computers got broken by last update
$DelaySettings.Add(

    [PSCustomObject]@{

        Name          = 'TwoDays'
        Collections   = ('Standard', 'NonCritical')
        ApprovalDelay = 2
    }
)

$DelaySettings.Add(

    [PSCustomObject]@{

        Name          = 'OneWeek'
        Collections   = ('Touchy', 'Critical')
        ApprovalDelay = 7
    }
)

$firstOfThisMonth = (Get-Date $Now -Day 1 )

switch($firstOfThisMonth.DayOfWeek) {

    "Sunday" {$thisPatchTuesday = $firstOfThisMonth.AddDays(9)}
    "Monday" {$thisPatchTuesday = $firstOfThisMonth.AddDays(8)}
    "Tuesday" {$thisPatchTuesday = $firstOfThisMonth.AddDays(7)}
    "Wednesday" {$thisPatchTuesday = $firstOfThisMonth.AddDays(13)}
    "Thursday" {$thisPatchTuesday = $firstOfThisMonth.AddDays(12)}
    "Friday" {$thisPatchTuesday = $firstOfThisMonth.AddDays(11)}
    "Saturday" {$thisPatchTuesday = $firstOfThisMonth.AddDays(10)}
}

if ($Now.Date -le $thisPatchTuesday.Date) {

    $patchTuesday = $thisPatchTuesday

}

else {

    $firstOfNextMonth = (Get-Date $Now -Day 1 -Month ((Get-Date $Now).AddMonths(1).Month))

    switch($firstOfNextMonth.DayOfWeek) {

        "Sunday" {$patchTuesday = $firstOfNextMonth.AddDays(9);Break}
        "Monday" {$patchTuesday = $firstOfNextMonth.AddDays(8); Break}
        "Tuesday" {$patchTuesday = $firstOfNextMonth.AddDays(7); Break}
        "Wednesday" {$patchTuesday = $firstOfNextMonth.AddDays(13); Break}
        "Thursday" {$patchTuesday = $firstOfNextMonth.AddDays(12); Break}
        "Friday" {$patchTuesday = $firstOfNextMonth.AddDays(11); Break}
        "Saturday" {$patchTuesday = $firstOfNextMonth.AddDays(10); Break}
    }
}

#added this condition to fix syncday calculation

if($Now.Date -le $thisPatchTuesday.Date.AddDays($SyncDelay)) {

    $SyncDay = (Get-Date -Date $thisPatchTuesday).AddDays($SyncDelay)
}

else {

    $SyncDay = (Get-Date -Date $patchTuesday).AddDays($SyncDelay)
}

#end fix of syncday calculation

switch($Now.Date) {

    $patchTuesday.Date {

        $Action = $true
        $Comments.Add("==> It's patch Tuesday!`n")
        $Comments.Add("Next Sync will happen in $((New-TimeSpan -Start $Now.Date -End $SyncDay.Date).days) days on $($SyncDay.ToLongDateString())`n")
    }

    $SyncDay.Date {

        $Action = $true
        $Comments.Add("Next Patch Tuesday is in $((New-TimeSpan -Start $Now.Date -End $thispatchTuesday.Date).days) days on $($thispatchTuesday.ToLongDateString())`n")
        $Comments.Add("==> It's sync day! - synching WSUS with Microsoft!`n")
        (Get-WsusServer).GetSubscription().StartSynchronization()
    }

    default {

        $Action = $false
        $Comments.Add("Next Patch Tuesday is in $((New-TimeSpan -Start $Now.Date -End $patchTuesday.Date).days) days on $($patchTuesday.ToLongDateString())`n")
        $Comments.Add("Next Sync will happen in $((New-TimeSpan -Start $Now.Date -End $SyncDay.Date).days) days on $($SyncDay.ToLongDateString())`n")
    }
}

# Getting this once now, rather than for each iteration....
$WSUS = Get-WsusServer @WSUSServerParams
$AllUpdates = $WSUS.GetUpdates()
$AllTargetGroups = $WSUS.GetComputerTargetGroups()
$NeededUpdates = Get-WsusUpdate -Approval Unapproved -Status FailedOrNeeded

foreach($Schedule in $DelaySettings) {

    # Added handling of different date conditions
    if($Now.Date -eq (Get-Date -Date $thisPatchTuesday).AddDays($SyncDelay).AddDays($Schedule.ApprovalDelay).Date) {

        foreach ($Group in $Schedule.Collections) {

            $Action = $true
            $Comments.Add('Case 1 - ')
            $Comments.Add("==> It is $($Schedule.ApprovalDelay) days after Sync Day. Approving updates for $($Group)`n")
            #$NeededUpdates | Approve-WsusUpdate -Action Install -TargetGroupName $Group -Verbose
        }  
    }

    elseif(($Now.Date -lt (Get-Date -Date $thisPatchTuesday).AddDays($SyncDelay).AddDays($Schedule.ApprovalDelay).Date) -and ($Now.Date -gt $thispatchTuesday)) {

        $Comments.Add('Case 2 - ')
        $Comments.Add("Next approval for the $($Schedule.Name) schedule will happen in $((New-TimeSpan -Start $Now.Date -End (((Get-Date -Date $thisPatchTuesday).AddDays($SyncDelay)).AddDays($Schedule.ApprovalDelay).Date)).days) days on $((((Get-Date -Date $thisPatchTuesday).AddDays($SyncDelay)).AddDays($Schedule.ApprovalDelay).Date).ToLongDateString())`n")
    }

    elseif(($Now.Date -gt $SyncDay.AddDays($Schedule.ApprovalDelay).Date)) {

        $Comments.Add('Case 3 - ')
        $Comments.Add("Next approval for the $($Schedule.Name) schedule will happen in $((New-TimeSpan -Start $Now.Date -End (((Get-Date -Date $thisPatchTuesday).AddDays($SyncDelay)).AddDays($Schedule.ApprovalDelay).Date)).days) days on $((((Get-Date -Date $thisPatchTuesday).AddDays($SyncDelay)).AddDays($Schedule.ApprovalDelay).Date).ToLongDateString())`n")
    }

    else {

        $Comments.Add('Case 4 - ')
        $Comments.Add("Next approval for the $($Schedule.Name) schedule will happen in $((New-TimeSpan -Start $Now.Date -End (((Get-Date -Date $PatchTuesday).AddDays($SyncDelay)).AddDays($Schedule.ApprovalDelay).Date)).days) days on $((((Get-Date -Date $PatchTuesday).AddDays($SyncDelay)).AddDays($Schedule.ApprovalDelay).Date).ToLongDateString())`n")
    }
}

if($Now.day -eq $WSUSCleanUpDay) {

    $Action = $true
    $Comments.Add("==> Today is WSUS monthly clean up day`n")
}

else {

    $Comments.Add("Next WSUS monthly clean up will happen in $((New-TimeSpan -Start $Now.Date -End $(Get-Date -Day $WSUSCleanUpDay -Month $($Now.Month + 1) -OutVariable datenextcleanup).Date).Days) days on $($datenextcleanup.ToLongDateString())`n")
}

$Comments

if(!$Action) {

    $Comments.Add("<i style='color:red'>No actions to be done today</i>`n")
}

$CommentsHTML = "<p style='color:blue'>" + $Comments.replace("`n", '<br>') + "</p>"

# I don't know if the GetTotalSummary() method will pull live data every time it is called
# If it does, then we do NOT need to get all of the target computer groups again, and
# the line below this can be commented out

$AllTargetGroups = $WSUS.GetComputerTargetGroups()

$PatchReport = $AllTargetGroups | ForEach {

    $Group = $_.Name

    $_.GetTotalSummary() | ForEach {

        [PSCustomObject] @{

            TargetGroup               = $Group
            Needed                    = ($_.NotInstalledCount + $_.DownloadedCount)
            "Installed/NotApplicable" = ($_.NotApplicableCount + $_.InstalledCount)
            NoStatus                  = $_.UnknownCount
            PendingReboot             = $_.InstalledPendingRebootCount
        }
    }
}

$params = @{

    'Encoding' = [System.Text.Encoding]::UTF8
    'To' = 'admin@yourdomain.com'
    'From' = 'wsus@yourdomain.com'
    'SmtpServer' = "smtpserver"
    'BodyAsHtml' = $true
    'Subject' = "WSUS Report"
    'Body' = (($CommentsHTML) + "<br>" + ($PatchReport | ConvertTo-Html | Out-String))  
}

Send-MailMessage @Params
