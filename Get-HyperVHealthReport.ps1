#requires -version 4
#requires -Modules Hyper-V

[CmdletBinding()]
param (
	
	[Parameter( Mandatory=$false)]
	[string[]]$Server=".",
    
    [Parameter( Mandatory=$false)]
    [switch]$SendEmail

	)

#region Begin Main Script
Begin {

    $smtpsettings = @{
	    To =  "foo@foo.com"
	    From = "bar@bar.com"
        Subject = "Hyper-V Health Report"
	    SmtpServer = "mailhost"
	    }

    $now = (Get-Date).ToShortDateString()

    #Common HTML head and styles
	$htmlhead="<html>
				<style>
				BODY{font-family: Arial; font-size: 8pt;}
				H1{font-size: 20px;}
				H2{font-size: 18px;}
				H3{font-size: 16px;}
				TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
				TH{border: 1px solid black; background: #dddddd; padding: 5px; color: #000000;}
				TD{border: 1px solid black; padding: 5px; }
				td.pass{background: #7FFF00;}
				td.warn{background: #FFE600;}
				td.fail{background: #FF0000; color: #ffffff;}
				td.info{background: #85D4FF;}
				</style>
				<body>
				<h1 align=""center"">Hyper-V Health Report</h1>
				<h3 align=""center"">Generated: $now</h3>"

    $htmltail = "</body>
			</html>"

    $spacer = "<br />"

    $report = @()
    
    Write-Verbose "Retrieving VM host details."
    $VMHosts = @(Get-VMHost $Server)

}

Process {

    foreach ($VMHost in $VMHosts)
    {
        $props = [ordered]@{
                            Name = $($VMHost.Name)
                            Domain = $($VMHost.FullyQualifiedDomainName)
                            Memory = "$("{0:0}" -f $($VMHost.MemoryCapacity /1GB))GB"
                            }
        
        $VMHostObject = New-Object -TypeName PSObject -Property $props

        Write-Verbose "Retrieving VM guest details for host: $($VMHost.Name)"
        $vms = Get-VM -ComputerName $VMHost.Name | Select-Object -Property Name,State,CPUUsage,MemoryAssigned,Uptime,Status

        $VMHostHtml = "<h3>Host: $($VMHost.Name.ToUpper())</h3>"
        $VMHostHtml += $VMHostObject | ConvertTo-Html -Fragment
        $VMHostHtml += $spacer
        $VMHostHtml += $vms | ConvertTo-Html -Fragment
        $VMHostHtml += $spacer

        $report += $VMHostHtml

    }
}

End {

    Write-Verbose "Generating report."
    
    $reportFile = "$PSScriptRoot\Hyper-V-Health-Report.html"

    $htmlreport = $htmlhead + $report + $htmltail

    $htmlreport | Out-File $reportFile -Encoding Utf8

    if ($SendEmail)
    {
        Write-Verbose "Sending email."
        Send-MailMessage @smtpsettings -Body $htmlreport -BodyAsHtml -Encoding ([System.Text.Encoding]::UTF8)
    }

    Write-Verbose "Finished."

}
#endregion Main Script