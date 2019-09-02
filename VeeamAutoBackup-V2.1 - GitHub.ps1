<# 
==============================================================================================================================
Author: Adam Ainsworth
Release Date: 02/09/2019
LinkedIn: https://uk.linkedin.com/in/adam-ainsworth-1ba448144?trk=people-guest_profile-result-card_result-card_full-click
Github: https://github.com/aainsworth-tech

This script expands upon the work carried out by Vladimir Eremin, Product Manager at Veeam. 

FREE VEEAM AUTOMATION SCRIPT
The purpose of this script is to provide free Veeam backup automation for all those poor admins who do not receive the budget they need.

There are several additions to the origional script written by Vladimir. This version will provide details for backups with the 'Warning' Status and format the 
email appropriately. I have additionally included retry functionality in case a backup session should fail (There is a setting for this in the options section). 
A sleep function has been added which will delay running of the email report code until a specified time, which allows all the reports to come through simultaneously. 
This is useful if you want to review all of your backup results at a specific time. If deploying across multiple sites, you can use the time that your last backup 
finishes as a guide for when to set the report time. E.g. Last site backup finishes at 11:30am, so set the report delay to "12:00:00 pm" (Use 12 hour clock and add PM or AM).

**NOTE** You may need to update line 124 depending on your hypervisor. If using Hyper-V, use Find-VBRHvEntity. For VMware ESXi use Find-VBRViEntity.

==============================================================================================================================
#>

$ErrorActionPreference = "SilentlyContinue"

#Add the Veeam PowerShell snap-in
Try
   {
    Asnp VeeamPSSnapin
   }
Catch
   {
   write-host "Failed to add Veeam Powershell Snapin." -ForegroundColor Red
   }

#===================================Options==========================================

    #VM name(s). Separate with Commas if multiple
    $VMNames = "VM Name Here"

    #Hypervisor
    $HostName = “Hypervisor Domain Name / IP”

    #Directory for backups
    $Directory = "Backup Directory"
    
    #Compression level
    $CompressionLevel = “4”
	
	#Maximum Retries
	$Retries = "2"
    
    #SMTP server
    $SMTPServer = “SMTP IP / Domain Name”

    #Send Backup Report?
    $EmailReport = $True

    #Backup Report Time
    $ReportTime = "12:00:00 pm"
    
    #Email FROM
    $EmailFrom = "Email Address Here"
    
    #EmailTo - SUCCESS 
    $EmailToSUCCESS = "Email Address Here"

    #EmailTo - FAILURE
    $EmailToFAILED = "Email Address Here"

    # Email subject
    $Date = Get-Date -Format g
    $EmailSubject = “Site / Server Name $Date"
    
    #=====================Email Subject formatting=====================

    $EmailSubjectSUCCESS = "Site / Server Name Backup SUCCESS $Date"

    $EmailSubjectWARNING = "Site / Server Name Backup WARNING $Date"

    $EmailSubjectFAILED = "Site / Server Name Backup FAILED $Date"


#=====================================End of Options============================================


#=====================================Email Styles==============================================

#Success HTML Style Sheet
$styleSUCCESS = "<style>BODY{font-family: Tahoma; font-size: 8pt; color: #4C607B;}"
$styleSUCCESS = $styleSUCCESS + "TABLE{border: 1px solid #A0A0A0; border-collapse: collapse;}"
$styleSUCCESS = $styleSUCCESS + "TH{border: 1px solid #A0A0A0; background: #00B238; color: #ffffff; padding: 5px;}"
$styleSUCCESS = $styleSUCCESS + "TD{border: 1px #000000; padding: 5px;}"
$styleSUCCESS = $styleSUCCESS + "</style>"

#Warning HTML Style Sheet
$styleWARNING = "<style>BODY{font-family: Tahoma; font-size: 8pt; color: #4C607B;}"
$styleWARNING = $styleWARNING + "TABLE{border: 1px solid #A0A0A0; border-collapse: collapse;}"
$styleWARNING = $styleWARNING + "TH{border: 1px solid #A0A0A0; background: #FFC400; color: #ffffff; padding: 5px;}"
$styleWARNING = $styleWARNING + "TD{border: 1px #000000; padding: 5px;}"
$styleWARNING = $styleWARNING + "</style>"

#Failed HTML Style Sheet
$styleFAILED = "<style>BODY{font-family: Tahoma; font-size: 8pt; color: #4C607B;}"
$styleFAILED = $styleFAILED + "TABLE{border: 1px solid #A0A0A0; border-collapse: collapse;}"
$styleFAILED = $styleFAILED + "TH{border: 1px solid #A0A0A0; background: #FB5A5A; color: #ffffff; padding: 5px;}"
$styleFAILED = $styleFAILED + "TD{border: 1px #000000; padding: 5px;}"
$styleFAILED = $styleFAILED + "</style>"


#=========Backup Section: Contacts hypervisor and starts backup session. Attempts backup a maximum of three times. Stops the loop when back is successful and moves on to email=========


$Retry = $True
[int]$Retrycount = "0"

Do { 
   Try {
        $Server = Get-VBRServer -name $HostName
        $mbody = @()
        foreach ($VMName in $VMNames)
        {
        $VM = Find-VBRHvEntity -Server $Server | Where {$_.Name -eq $VMName}
        $fullpath = join-path -path $Directory -childpath $VMName 
        $ZIPSession = Start-VBRZip -Entity $VM -Folder $fullpath -Compression $CompressionLevel
        $Retry = $False
        }
    }

 Catch {
        If ($Retrycount -gt $Retries) {
        Write-Host "Backup failed after $Retrycount retries." -ForegroundColor Red
		Start-sleep -seconds 3
        $Retry = $False
        }
        Else {
        Write-host "Backup failed. Retrying in 10 seconds..." -ForegroundColor Magenta
        Start-sleep -seconds 10
        $Retry = $Retrycount + 1
        }
    }
}  
While ($Retry -eq $True)


#==================Custom sleep function: Stops email report coming through until specified time=====================
   
   
 Function sleep-until($future_time) { 
    if ([String]$future_time -as [DateTime]) { 
        if ($(get-date $future_time) -gt $(get-date)) { 
            $sec = [system.math]::ceiling($($(get-date $future_time) - $(get-date)).totalseconds) 
            start-sleep -seconds $sec 
        } 
        else { 
            write-host "You must specify a date/time in the future" 
            return 
        } 
    } 
    else { 
        write-host "Incorrect date/time format" 
    } 
}

#Set report time
Sleep-Until $ReportTime
  

#============================Email Report===============================


    $TaskSessions = $ZIPSession.GetTaskSessions().logger.getlog().updatedrecords
    $WarningSessions = $ZIPsession | where {$_.Result -eq "Warning"}
    $FailedSessions =  $TaskSessions | where {$_.Status -eq "EFailed"} 
    
  If ($FailedSessions -ne $Null) {
    $EmailTo = $EmailToFAILED
    $EmailSubject = $EmailSubjectFAILED
    $MesssagyBody = ($ZIPSession | Select-Object @{n="Name";e={($_.name).Substring(0, $_.name.LastIndexOf("("))}} ,@{n="Start Time";e={$_.CreationTime}},@{n="End Time";e={$_.EndTime}},Result,@{n="Details";e={$FailedSessions.Title}})
    $style = $styleFAILED
    $Priority = "High"
    }    
  ElseIf ($WarningSessions -ne $Null) {
    $EmailTo = $EmailToSUCCESS
    $EmailSubject = $EmailSubjectWARNING
    $MesssagyBody = ($ZIPSession | Select-Object @{n="Name";e={($_.name).Substring(0, $_.name.LastIndexOf("("))}} ,@{n="Start Time";e={$_.CreationTime}},@{n="End Time";e={$_.EndTime}},Result,@{n="Details";e={$ZIPSession.GetDetails()}})
    $style = $styleWARNING
    $Priority = "High"
    } 
  Else {
    $EmailTo = $EmailToSUCCESS
    $EmailSubject = $EmailSubjectSUCCESS
    $MesssagyBody = ($ZIPSession | Select-Object @{n="Name";e={($_.name).Substring(0, $_.name.LastIndexOf("("))}} ,@{n="Start Time";e={$_.CreationTime}},@{n="End Time";e={$_.EndTime}},Result,@{n="Details";e={($TaskSessions | sort creationtime -Descending | select -first 1).Title}})
    $style = $styleSUCCESS
    $Priority = "Normal"
    }


  If ($EmailReport) {
    foreach ($EmailTo in $EmailTo){
      $Message = New-Object System.Net.Mail.MailMessage $EmailFrom, $EmailTo
      $Message.Subject = $EmailSubject 
      $Message.Priority = $Priority
      $Message.IsBodyHTML = $True
      $message.Body = $MesssagyBody | ConvertTo-Html -head $style | Out-String 
      $SMTP = New-Object Net.Mail.SmtpClient($SMTPServer)
      $SMTP.Send($Message)
      }
   }