#author: happysysadm
$DelaySettings = New-Object System.Collections.ArrayList

$DelaySettings.Add(

    [PSCustomObject]@{

        Name          = 'Immediate'
        WsusGroup   = 'Standard'
        ApprovalDelay = 1
    }
)

$DelaySettings.Add(

    [PSCustomObject]@{

        Name          = 'OneWeek'
        WsusGroup   = 'Touchy','Critical'
        ApprovalDelay = 7
    }
)

$DelaySettings | ConvertTo-Json | Out-File ApprovalDelaySettings.json -Force
Get-Content ApprovalDelaySettings.json | ConvertFrom-Json
