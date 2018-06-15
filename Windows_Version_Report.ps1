#WCA 2018 | Malinda 
#
#
#Get the latest Build #
$URI = "https://pureinfotech.com/windows-10-version-release-history/"
$html = Invoke-WebRequest $URI
$links = $html.links | where {$_.innerhtml -like "17*"}
$latestBuild = $links[0].innerHTML
#
$array01=Get-ADComputer -Filter {OperatingSystem -eq "Windows 10 Pro"} -Property * | Where-Object {($_.OperatingSystemVersion -notlike "*16299*") -and ($_.OperatingSystemVersion -notlike "*17*") -and ($_.OperatingSystemVersion -notlike "*18*") }
[string]$ReportTitle = "TBL Windows Update Compliance Report"
[string]$ReportSmallTitle = "Workstations Below Windows 10 build 16299"
[string]$ReportFooter01 = "WCA 2018"
#
#[string]$ReportFileName =  "WSReport_$(get-date).html"
#
#HTML_Formatting
$convertParams = @{ 
 PreContent = "<H1>$($ReportTitle)</H1>","<H3>$($latestBuild)</H3>","<H3>$($ReportSmallTitle)</H3>"
 PostContent = "<p class='footer'>$(get-date)</p>", "<p class='footer'>$ReportFooter01</p>"
 head = @"
 <Title>$ReportTitle</Title>
<style>
body { background-color:#E5E4E2;
       font-family:calibri;
       font-size:10pt; }
td, th { border:0px solid black; 
         border-collapse:collapse;
         white-space:pre; }
th { color:white;
     background-color:black; }
table, tr, td, th { padding: 2px; margin: 0px ;white-space:pre; }
tr:nth-child(odd) {background-color: lightgray}
table { width:95%;margin-left:5px; margin-bottom:20px;}
h2 {
 font-family:calibri;
 color:#6D7B8D;
}
.footer 
{ color:darkred; 
  margin-left:10px; 
  font-family:calibri;
  font-size:8pt;
  font-style:italic;
}
</style>
"@
}
#
#SMTP and Send-mail configs 
#
$array01 | ConvertTo-Html @convertParams -Property Name,OperatingSystem,OperatingSystemVersion,OperatingSystem,Created,LastLogonDate | Out-File -FilePath C:\temp\WSReport.html
#
$EmailSubject = "TBL Windows Update Compliance Report_$(get-date -format 'dd-MMMM-yyyy')"
$EmailRecipients = "alerts@wcatech.com","mrathnayake@wcatech.com","envmon@broadway.org","NFreeman@broadway.org"
#$EmailRecipients = "mrathnayake@wcatech.com","envmon@broadway.org"
$EmailFrom = "envmon@broadway.org"
$EmailSMTPServer = 'broadway-org.mail.protection.outlook.com'
$EmailEncoding = [System.Text.Encoding]::UTF7 #UTF7,UTF8,ASCII
$Attachment = "C:\temp\WSReport.html"
#
Send-MailMessage -To $EmailRecipients -Subject $EmailSubject -From $EmailFrom -Body $EmailText -SmtpServer $EmailSMTPServer -Encoding $EmailEncoding -Attachments $Attachment
#