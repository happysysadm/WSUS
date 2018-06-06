$action = $false

$now = Get-Date

$comments = "Today is $($now.ToLongDateString())`n"

$d0 = Get-Date -Day 1 -Month $($now.Month) -Year $now.Year

switch ($d0.DayOfWeek){

        "Sunday"    {$patchTuesday0 = $d0.AddDays(9); break}

        "Monday"    {$patchTuesday0 = $d0.AddDays(8); break}

        "Tuesday"   {$patchTuesday0 = $d0.AddDays(7); break}

        "Wednesday" {$patchTuesday0 = $d0.AddDays(13); break}

        "Thursday"  {$patchTuesday0 = $d0.AddDays(12); break}

        "Friday"    {$patchTuesday0 = $d0.AddDays(11); break}

        "Saturday"  {$patchTuesday0 = $d0.AddDays(10); break}

     }

$d1 = Get-Date -Day 1 -Month $($now.Month + 1) -Year $now.Year

switch ($d1.DayOfWeek){

        "Sunday"    {$patchTuesday1 = $d1.AddDays(9); break}

        "Monday"    {$patchTuesday1 = $d1.AddDays(8); break}

        "Tuesday"   {$patchTuesday1 = $d1.AddDays(7); break}

        "Wednesday" {$patchTuesday1 = $d1.AddDays(13); break}

        "Thursday"  {$patchTuesday1 = $d1.AddDays(12); break}

        "Friday"    {$patchTuesday1 = $d1.AddDays(11); break}

        "Saturday"  {$patchTuesday1 = $d1.AddDays(10); break}

     }

if($now.date -le $patchTuesday0.date){

    $patchTuesday = $patchTuesday0}else{$patchTuesday = $patchTuesday1
    
    }

$d0 = Get-Date -Day 1 -Month $($now.Month) -Year $now.Year

switch ($d0.DayOfWeek){

        "Sunday"    {$FourthMonday0 = $d0.AddDays(22); break}

        "Monday"    {$FourthMonday0 = $d0.AddDays(21); break}

        "Tuesday"   {$FourthMonday0 = $d0.AddDays(20); break}

        "Wednesday" {$FourthMonday0 = $d0.AddDays(26); break}

        "Thursday"  {$FourthMonday0 = $d0.AddDays(25); break}

        "Friday"    {$FourthMonday0 = $d0.AddDays(24); break}

        "Saturday"  {$FourthMonday0 = $d0.AddDays(23); break}

     }

    
$d1 = Get-Date -Day 1 -Month $($now.Month + 1) -Year $now.Year

switch ($d1.DayOfWeek){

        "Sunday"    {$FourthMonday1 = $d1.AddDays(22); break}

        "Monday"    {$FourthMonday1 = $d1.AddDays(21); break}

        "Tuesday"   {$FourthMonday1 = $d1.AddDays(20); break}

        "Wednesday" {$FourthMonday1 = $d1.AddDays(26); break}

        "Thursday"  {$FourthMonday1 = $d1.AddDays(25); break}

        "Friday"    {$FourthMonday1 = $d1.AddDays(24); break}

        "Saturday"  {$FourthMonday1 = $d1.AddDays(23); break}

     }

if($now.date -le $FourthMonday0.date){

    $FourthMonday = $FourthMonday0}else{$FourthMonday= $FourthMonday1
    
    }

if($now.date -le $FourthMonday0.adddays(1).date){

    $StandardApprovalDay = $FourthMonday0.AddDays(1)}else{$StandardApprovalDay= $FourthMonday1.AddDays(1)
    
    }

if($now.date -le $FourthMonday0.adddays(1).date){

    $CriticalApprovalDay = $FourthMonday0.AddDays(7)}else{$CriticalApprovalDay= $FourthMonday1.AddDays(7)
    
    }

if($now.date -eq $PatchTuesday.date){

    $comments += "==> It's patch Tuesday!`n"

    $action = $true

    }

else {

    $comments += "Next Patch Tuesday is in $((New-TimeSpan -Start $now.date -End $patchTuesday.date).days) days on $($patchTuesday.ToLongDateString())`n"
    
    }

if($now.date -eq $FourthMonday.date){

    $comments += "==> It's fourth monday of the month - synching WSUS with Microsoft!`n"

    $action = $true

    $startTime = (get-date -f dd-MM-yyyy)

    (Get-WsusServer).GetSubscription().StartSynchronization()

    }

else {

    $comments += "Next Sync will happen in $((New-TimeSpan -Start $now.date -End $FourthMonday.date).days) days on $($FourthMonday.ToLongDateString())`n"
    
    }

if($now.date -eq $StandardApprovalDay.date){

    $comments += "==> It's the day after fourth monday of the month - approving for Standard servers`n"

    $action = $true

    $wsus = Get-WsusServer

    $allupdates = $wsus.GetUpdates() 

    $alltargetgroups = $wsus.GetComputerTargetGroups()

    $computergroups = ($alltargetgroups | ? name -match 'Standard').name

    $computergroups | % {

        Get-WsusUpdate -Approval Unapproved -Status FailedOrNeeded | Approve-WsusUpdate -Action Install -TargetGroupName $_ –Verbose

        }

    $startTime = (Get-Date -f dd-MM-yyyy)

    }

else {

    $comments += "Next approval for Standard servers will happen in $((New-TimeSpan -Start $now.date -End $StandardApprovalDay.Date).days) days on $($StandardApprovalDay.ToLongDateString())`n"
    
    }

if($now.date -eq $CriticalApprovalDay.date){

    $comments += "==> It's the 7th day after fourth monday of the month - approving for User-Touchy and Mission-Critical servers`n"

    $action = $true

    $wsus = Get-WsusServer

    $allupdates = $wsus.GetUpdates() 

    $alltargetgroups = $wsus.GetComputerTargetGroups()

    $computergroups = ($alltargetgroups | ? name -match 'touchy|critical').name

    $computergroups | % {

        Get-WsusUpdate -Approval Unapproved -Status FailedOrNeeded | Approve-WsusUpdate -Action Install -TargetGroupName $_ –Verbose

        }

    $startTime = (get-date -f dd-MM-yyyy)

    }

else {

    $comments += "Next approval for User-Touchy and Mission-Critical servers will happen in $((New-TimeSpan -Start $now.date -End $CriticalApprovalDay.date).days) days on $($CriticalApprovalDay.ToLongDateString())`n"
    
    }

if($now.day -eq 7){

    $comments += "==> Today is WSUS monthly clean up day`n";$action = $true

    }

else{

    $comments += "Next WSUS monthly clean up will happen in $((New-TimeSpan -Start $now.date -End $(Get-Date -Day 7 -Month $($now.Month + 1) -Year $now.Year -OutVariable datenextcleanup).Date).Days) days on $($datenextcleanup.ToLongDateString())`n"

    }

$comments

if(!$action){$comments += "<i style='color:red'>No actions to be done today</i>`n"}

$commentshtml = "<p style='color:blue'>" + $comments.replace("`n",'<br>') + "</p>"

$wsus = Get-WsusServer

$alltargetgroups = $wsus.GetComputerTargetGroups()

$patchreport = $alltargetgroups | ForEach {

    $Group = $_.Name

    $_.GetTotalSummary() | ForEach {

        [pscustomobject]@{

            TargetGroup = $Group

            Needed = ($_.NotInstalledCount + $_.DownloadedCount)

            "Installed/NotApplicable" = ($_.NotApplicableCount + $_.InstalledCount)

            NoStatus = $_.UnknownCount

            PendingReboot = $_.InstalledPendingRebootCount

        }

    }

}

$params = @{
    
        'encoding'=[System.Text.Encoding]::UTF8
		
        'To' = 'recipient@domain.com'

        'From' = 'sender@domain.com'

        'SmtpServer' = "smtphost"

        'BodyAsHtml' = $true

        'Subject' = "WSUS - Patch Report"
        
        'Body' = (($commentshtml) + "<br>" + ($patchreport | ConvertTo-Html | Out-String))
   
        }

Send-MailMessage @params