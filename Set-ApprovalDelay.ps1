$DelaySettings = @()

$DelaySettings += [pscustomobject]@{

    Name          = 'Immediate'

    WsusGroup   = 'Standard'

    ApprovalDelay = 1

}

$DelaySettings += [pscustomobject]@{

    Name          = 'OneWeek'

    WsusGroup   = 'Touchy','Critical'

    ApprovalDelay = 7

}


$DelaySettings | ConvertTo-Json  | Out-File approvaldelaysettings.json -Force

Get-Content approvaldelaysettings.json | ConvertFrom-Json