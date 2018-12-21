<#
.SYNOPSIS
	Modify Exchange parameters to help resolve issue with EventID 3429 appearing in Veritas Enterprise Vault logs.

.DESCRIPTION
	Script creates a remote Powershell session to your Exchange server to modify Exchange attributes 'ProhibitSendReceiveQuota', 'ProhibitSendReceiveQuota' 
	and 'UseDatabaseQuotaDefaults' for each unique user associated with EventID-3429.
	A 'RUN NOW' has to be performed on the Archive Task in EV for mailboxes which were unable to synchronise due to mailbox is over quota and desktop limit.
	
	The following information is relevant for running this script:
		Log Name:		Veritas Enterprise Vault
		Source:			Enterprise Vault
		EventID:		3429
		Task Category:	Archive Task

	*** This has been tested in an environment using Enterprise Vault 12.3 with Exchange 2013 ***

.PARAMETER Log
	Optional log creation, but good for auditing and troubleshooting purposes.

.PARAMETER SendMail
	Sends an HTML report via email using the SMTP configuration from the EV_Oldaccounts-Settings.xml file includes attachments.

.EXAMPLE 
	./EV_oldaccounts.ps1 -Log  - This will run the script and provide an audit file at the end of the process

.NOTES
	Version:        1.0
	Author:         Jason McColl
	Email:			jason.mccoll@outlook.com
	Creation Date:  18/12/2018
	Thanks:			Jurgis Primagovas (@primaju), my friend for fixing my bad coding 
	
MIT License

Copyright (c) 2018 Jason M McColl

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated
documentation files (the "Software"), to deal in the Software without restriction, including without limitation
the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software,
and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
The above copyright notice and this permission notice shall be included in all copies or substantial portions of
the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO
THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

Change Log
	V1.00, 18/12/2018 - Initial version

#>
#requires -version 2

[CmdletBinding()]
param (
	[Parameter( Mandatory=$false)]
	[switch]$Log,
	
	[Parameter( Mandatory=$false)]
	[switch]$SendMail
)

#-----------------------------------------------------------[Settings]------------------------------------------------------------

#Don't change these, or you will break the script 
$Credentials = Get-Credential
$now = Get-Date													#Used for timestamps
$date = $now.ToShortDateString()								#Short date format for email message subject
$myDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ErrorActionPreference = 'SilentlyContinue'
$StartDate = "01/01/1900"
$EndDate = (Get-Date)
$pattern = ".*(?= on Ex)"										#REGEX pattern 

#Modify these paramters
$mailserver = "Exchange.name.here"								#Enter the FQDN of your Exchange server 
$reportemailsubject = "Modify Exchange attributes for accounts which report EventID:3429 in EV - $date"
$quotasize = "10737418240"										#This size is in bytes, change this value for your organisation
$AuditLog = "$myDir\EventID-3429-Audit-$(get-date -f yyyy-MM-dd-HHmmss).log"
$mbxprocessed = "$myDir\mbxlist.txt";						

#Create Remote PowerShell session to Exchange server
$tmpstring = "Creating a remote PowerShell session to the defined Exchange server '$mailserver'"
if ($Log) {Write-Logfile $tmpstring}
Get-PSSession | where {$_.ConfigurationName -eq "Microsoft.Exchange"} | Remove-PSSession
$Session=New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$mailserver/PowerShell/" -Authentication Kerberos -Credential $Credentials
Import-PSSession $Session

### Import EV_3429.xml config file - Modify this file to reflect your SMTP settings and any Exclusions
[xml]$ConfigFile = Get-Content "$myDir\EV_3429.xml"  # Change to the location where you store these files

### Email settings - Modify entries in the EV_Oldaccounts-Settings.xml
$smtpsettings = @{
    To = $ConfigFile.Settings.EmailSettings.MailTo
    From = $ConfigFile.Settings.EmailSettings.MailFrom
    Subject = $reportemailsubject
    SmtpServer = $ConfigFile.Settings.EmailSettings.SMTPServer
}

#-----------------------------------------------------------[Functions]------------------------------------------------------------

### This function is used to write the log file if -Log is used
Function Write-Logfile()
{
	param( $logentry )
	$timestamp = Get-Date -DisplayHint Time
	"$timestamp $logentry" | Out-File $AuditLog -Append
}

#-----------------------------------------------------------[Initialisation]------------------------------------------------------------

### Logfile Strings
$logstring0 = "===================================================="
$logstring1 = "Processing of accounts reporting EventID 3429 on EV server"
$initstring0 = "Initializing..."

### Log file is overwritten each time the script is ran to prevent many large log files
if ($Log)
{
	$timestamp = Get-Date -DisplayHint Time
	"$timestamp $logstring0" | Out-File $AuditLog
	Write-Logfile $logstring1
	Write-Logfile $logstring0
	Write-Logfile $initstring0
}

### Obtain a list of identities from Event Log who's mailbox is over quota

$tmpstring = "Generting list of users from all EventID-3429 from Veritas Enteprise Vault log"
if ($Log) {Write-Logfile $tmpstring}
$EV_3429 = Get-WinEvent -FilterHashTable @{LogName = "Veritas Enterprise Vault"; ID = "3429"; StartTime = $StartDate; EndTime = $EndDate}
$usermbxs = $EV_3429 | %{ [regex]::matches($_.Properties[0].Value,$pattern).value[0]} | sort -Unique
$mbxlist = $usermbxs | Out-File -FilePath $mbxprocessed


### Attempting to install AD Module

foreach ($mbx in $usermbxs)
{
	$status = (get-mailbox $mbx)
	if ($status.UseDatabaseQuotaDefaults -eq 'False')
	{
	$tmpstring = "Quotas changed for $($status.UserPrincipalName)"
	if ($Log) {Write-Logfile $tmpstring}
	Set-Mailbox $mbx -ProhibitSendReceiveQuota $quotasize -ProhibitSendQuota $quotasize
	}
	else
	{
	$tmpstring = "$($status.UserPrincipalName) required 'UseDatabaseQuotaDefaults', 'ProhibitSendReceiveQuota' and 'ProhibitSendQuota' to be modified"
	if ($Log) {Write-Logfile $tmpstring}
	sleep 1
	Set-Mailbox $mbx -UseDatabaseQuotaDefaults $False -ProhibitSendReceiveQuota $quotasize -ProhibitSendQuota $quotasize 
	}
}

#-----------------------------------------------------------[Email Configuration]------------------------------------------------------------

### Email generation
	
	$report = "Total Mailboxes Checked:" + $usermbxs.count 
	
if ($SendMail)
{
    $reporthtml = $report | ConvertTo-Html -Fragment
	$htmlhead="<html>
				<style>
				BODY{font-family: Arial; font-size: 8pt;}
				H1{font-size: 22px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
				H2{font-size: 18px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
				H3{font-size: 16px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
				TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
				TH{border: 1px solid #969595; background: #dddddd; padding: 5px; color: #000000;}
				TD{border: 1px solid #969595; padding: 5px; }
				td.pass{background: #B7EB83;}
				td.warn{background: #FFF275;}
				td.fail{background: #FF2626; color: #ffffff;}
				td.info{background: #85D4FF;}
				</style>
				<body>
                <p>Report of mailboxes requiring changes to allow archiving on $date. </p>
				<p>CSV version of report attached to this email.</p>"
		
	$htmltail = "</body></html>"	

	$htmlreport = $htmlhead + $reporthtml + $htmltail
    Send-MailMessage @smtpsettings -Body $htmlreport -BodyAsHtml -Encoding ([System.Text.Encoding]::UTF8) -Attachments @($AuditLog,$mbxprocessed)
	### Only delete log files when they have been sent in an email
	Write-Host "Removing Log files" -ForegroundColor Yellow
    Remove-Item $AuditLog;
	Remove-Item $mbxprocessed;
	#Remove-Item $myDir\EventID-3429.txt

	Write-Host "--Successfully removed the log files for this session--" -ForegroundColor Green
	
}

#-----------------------------------------------------------[House Keeping]------------------------------------------------------------

### Cleaning up the mess I made
Write-Host "Removing Microsoft Exchange PowerShell Session" -ForegroundColor Yellow
Remove-PSSession $Session
Write-Host "--Successfully ended the Microsoft Exchange PowerShell Session--" -ForegroundColor Green


