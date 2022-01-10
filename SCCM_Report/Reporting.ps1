Import-Module $PSScriptRoot\Out-PieChart.psm1
& "$PSScriptRoot\ISEConnect.ps1" 

$SCCMSiteCode = New-Object –ComObject “Microsoft.SMS.Client” 
$day=[datetime]::Today.DayOfWeek 
$bodypath= "c:\temp\body.txt" 
$date= get-date

if(Test-Path $bodypath){ 
    Remove-Item -Path $bodypath -Force 
} 

if(Test-Path C:\temp\*.png){ 
    foreach($img in (gci c:\temp | where {$_.Name -like "*status*.png"}))
    { 
        Remove-Item -Path $img.FullName -Force 

    } 
} 

$collection= Get-CMCollection -Name "*$day*" | select Name,CollectionId 
$deployments=Get-CMSoftwareUpdateDeployment | where {$_.EnforcementDeadline -eq $($date.ToString("MM/dd/yy"))} 

if(-not $deployments){ 
exit 
} 

foreach($deployment in $deployments){ 


    $query = @{ 
        Namespace = "root\SMS\site_$($SCCMSiteCode.GetAssignedSite())" 
        ClassName = 'SMS_SUMDeploymentAssetDetails' 
        Filter = "CollectionId like '$($deployment.TargetCollectionID)'" 
    } 

    $status = Get-CimInstance @query | select DeviceName,StatusType,StatusTime,StatusDescription,CollectionName,AssignmentName 

        foreach($p in $status)
            { 

                if($p.statustype -eq 1){$p.statustype = "Success"} 
                if($p.statustype -eq 2){$p.statustype = "InProgress"} 
                if($p.statustype -eq 3){$p.statustype = "RequirementsNotMet"} 
                if($p.statustype -eq 4){$p.statustype = "Unknown"} 
                if($p.statustype -eq 5){$p.statustype = "Error"} 

            } 

  

  

  

    $status | Group-Object -Property Statustype -NoElement | Out-PieChart -PieChartTitle "Status" -Pie3D -ChartWidth 400 -ChartHeight 400 -saveImage C:\temp\status_$($deployment.TargetCollectionID).png 

  

    $Success=@() 
    $InProgress=@() 
    $Requirements=@() 
    $Unknown=@() 
    $ErrorStatus=@() 

  

  

  

    foreach($machine in $status){ 

  

        if($machine.statustype -eq "Success"){ 

            $Success += $machine 

        } 

        if($machine.statustype -eq "InProgess"){ 

            $InProgress += $machine 

        } 

        if($machine.statustype -eq "RequirementsNotMet"){ 

            $Requirements += $machine 

        } 

        if($machine.statustype -eq "Unknown"){ 

            $Unknown+= $machine 

        } 

        if($machine.statustype -eq "Error"){ 

            $ErrorStatus += $machine 

        } 

    } 

  

  

    $board = [PSCustomObject]@{ 

        Success = ($Success |Measure-Object).Count 

        InProgress       = ($InProgress|Measure-Object).Count 

        RequirementsNotMet   = ($Requirements|Measure-Object).Count 

        Unknown    = ($Unknown|Measure-Object).Count 

        Error = ($ErrorStatus|Measure-Object).Count 

    } 

  

  

    $board 
    $bodyhtml="" 
    $bodyall="" 

        if(($status | where {$_.statustype -NotLike "Success"}) -eq $null){ 

            $body="" 

            $bodyhtml=
            @" 

            <p>$($machine.collectionName) <p> 

            <p>Success :<span style="color: #99cc00;"> $($board.Success)</span><br/>InProgress : $($board.InProgress)<br />RequirementsNotMet : $($board.RequirementsNotMet)<br />Unknown : $($board.Unknown)<br />Error : <span style="color: #ff0000;">$($board.Error)<br/></span><img src='cid:status_$($deployment.TargetCollectionID).png'></p> 
"@ 

$bodyhtml >> c:\temp\body.txt 

            }

             else
                { 

                    foreach($device in ($status | where {$_.statustype -NotLike "Success"})){ 

                    if($device -eq $null){} 

                    $body=@" 
                    <tr> 

                    <td style="width: 33%;">$($device.DeviceName)</td> 

                    <td style="width: 33%;">$($device.statustype)</td> 

                    <td style="width: 33%;">$($device.StatusDescription)</td> 

                    </tr> 

"@ 

  

                    $bodyall += $body 
                } 

$bodyhtml=@" 
<p>$($machine.collectionName) <p> 

<p>Success :<span style="color: #99cc00;"> $($board.Success)</span><br/>InProgress : $($board.InProgress)<br />RequirementsNotMet : $($board.RequirementsNotMet)<br />Unknown : $($board.Unknown)<br />Error : <span style="color: #ff0000;">$($board.Error)<br/></span><img src='cid:status_$($deployment.TargetCollectionID).png'></p> 

<table style="border-collapse: collapse; width: 100%; height: 18px;" border="1"> 

<tbody> 

<tr style="height: 18px;"> 

<td style="width: 33%; height: 18px;">ComputerName</td> 

<td style="width: 33%; height: 18px;">Status</td> 

<td style="width: 33%; height: 18px;">StatusDescription</td> 

$bodyall 

</tr> 

</tbody> 

</table> 
"@ 
$bodyhtml >> c:\temp\body.txt 

} 

} 


$html = gc c:\temp\body.txt 


if($html -eq $null){} 

else{ 

Send-MailMessage -From * -To * -BodyAsHtml  ($html | out-string ) -Subject "[$($env:COMPUTERNAME)] Deployment of $(get-date -f MM/dd/yyyy)" -Attachments $(gci C:\temp\*status*.png).FullName -SmtpServer * 

} 

 