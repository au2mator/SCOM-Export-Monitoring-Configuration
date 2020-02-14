
#########
# au2mator PS Services
# New Service
# SCOM - Export Monitoring Configuration
# v 1.0 Initial Release
# Init Release: 03.02.2020
# Last Update: 03.02.2020
# Code Template V 1.1
# URL: https://au2mator.com/export-monitoring-configuration-scom-self-service-with-au2mator/
# Github: https://github.com/au2mator/SCOM-Export-Monitoring-Configuration
#
# Based on https://gallery.technet.microsoft.com/scriptcenter/ExportEffectiveMonitoringCo-05d58912?fbclid=IwAR1dKU7eXTHy6FAu0yU1CJBtRGOBBJA6yQkg4XslaiwN1eEFE2Yicg5Y0ek
# Tyson Paul
#
#################


#region InputParamaters
##Question in au2mator
Param 
( 
    [Parameter(Mandatory = $false)]
    [string]$c_OptionType, 

    [Parameter(Mandatory = $false)]
    [string]$c_ComputerName, 

    [Parameter(Mandatory = $false)]
    [string]$c_SCOMGroupName, 

    [Parameter(Mandatory = $false)]
    #[ValidatePattern("[a-f0-9][a-f0-9][a-f0-9][a-f0-9][a-f0-9][a-f0-9][a-f0-9][a-f0-9]-[a-f0-9][a-f0-9][a-f0-9][a-f0-9]-[a-f0-9][a-f0-9][a-f0-9][a-f0-9]-[a-f0-9][a-f0-9][a-f0-9][a-f0-9]-[a-f0-9][a-f0-9][a-f0-9][a-f0-9][a-f0-9][a-f0-9][a-f0-9][a-f0-9][a-f0-9][a-f0-9][a-f0-9][a-f0-9]")] 
    [String]$c_ID, 

    [Parameter(Mandatory = $false)]
    [string]$c_OptionExport, 

    [Parameter(Mandatory = $false)]
    [string]$c_ExportPath, 

    [Parameter(Mandatory = $false)]
    [string]$c_ExportMail, 

    ## au2mator Initialize Data
    [parameter(Mandatory = $true)] 
    [String]$InitiatedBy, 

    [parameter(Mandatory = $true)] 
    [String]$RequestId, 

    [parameter(Mandatory = $true)] 
    [String]$Service, 

    [parameter(Mandatory = $true)] 
    [String]$TargetUserId
    )
#endregion  InputParamaters

#region Variables
##Script Handling
$DoImportPSSession = $false
$ErrorCount = 0


## Environment
[string]$SCOMServer = 'Demo01.au2mator.local'
[string]$LogPath = "C:\_SCOworkingDir\TFS\PS-Services\SCOM - Export Monitoring Configuration"
[string]$LogfileName = "Export Monitoring Configuration"

$TargetFolder="C:\_SCOworkingDir\TFS\PS-Services\SCOM - Export Monitoring Configuration\Export"

## au2mator Settings
[string]$PortalURL = "http://demo01.au2mator.local"
[string]$au2matorDBServer = "demo01"
[string]$au2matorDBName = "au2mator"

## Control Mail
$SendMailToInitiatedByUser = $true #Send a Mail after Service is completed
$SendMailToTargetUser = $true #Send Mail to Target User after Service is completed

## SMTP Settings
$SMTPServer = "smtp.office365.com"
$SMTPUser = "mail@au2mator.com"
$SMTPPassword = "Password1"
$SMPTAuthentication = $true #When True, User and Password needed
$EnableSSLforSMTP = $true
$SMTPSender = "mail@au2mator.com"
#endregion Variables


#region Functions

Function ConnectToDB {
    # define parameters
    param(
        [string]
        $servername,
        [string]
        $database
    )
    # create connection and save it as global variable
    $global:Connection = New-Object System.Data.SQLClient.SQLConnection
    $Connection.ConnectionString = "server='$servername';database='$database';trusted_connection=false; integrated security='true'"
    $Connection.Open()
    Write-Verbose 'Connection established'
}

Function ExecuteSqlQuery {
    # define parameters
    param(
      
        [string]
        $sqlquery
     
    )
     
    Begin {
        If (!$Connection) {
            Throw "No connection to the database detected. Run command ConnectToDB first."
        }
        elseif ($Connection.State -eq 'Closed') {
            Write-Verbose 'Connection to the database is closed. Re-opening connection...'
            try {
                # if connection was closed (by an error in the previous script) then try reopen it for this query
                $Connection.Open()
            }
            catch {
                Write-Verbose "Error re-opening connection. Removing connection variable."
                Remove-Variable -Scope Global -Name Connection
                throw "Unable to re-open connection to the database. Please reconnect using the ConnectToDB commandlet. Error is $($_.exception)."
            }
        }
    }
     
    Process {
        #$Command = New-Object System.Data.SQLClient.SQLCommand
        $command = $Connection.CreateCommand()
        $command.CommandText = $sqlquery
     
        Write-Verbose "Running SQL query '$sqlquery'"
        try {
            $result = $command.ExecuteReader()      
        }
        catch {
            $Connection.Close()
        }
        $Datatable = New-Object "System.Data.Datatable"
        $Datatable.Load($result)
        return $Datatable         
    }
    End {
        Write-Verbose "Finished running SQL query."
    }
}

Function Write-au2matorLog {
    [CmdletBinding()]
    param
    (
        [ValidateSet('DEBUG', 'INFO', 'WARNING', 'ERROR')]
        [string]$Type,
        [string]$Text
    )
       
    # Set logging path
    if (!(Test-Path -Path $logPath)) {
        try {
            $null = New-Item -Path $logPath -ItemType Directory
            Write-Verbose ("Path: ""{0}"" was created." -f $logPath)
        }
        catch {
            Write-Verbose ("Path: ""{0}"" couldn't be created." -f $logPath)
        }
    }
    else {
        Write-Verbose ("Path: ""{0}"" already exists." -f $logPath)
    }
    [string]$logFile = '{0}\{1}_{2}.log' -f $logPath, $(Get-Date -Format 'yyyyMMdd'), $LogfileName
    $logEntry = '{0}: <{1}> <{2}> <{3}> {4}' -f $(Get-Date -Format dd.MM.yyyy-HH:mm:ss), $Type, $RequestId, $Service, $Text
    Add-Content -Path $logFile -Value $logEntry
}

Function Get-UserInput ($RequestID) {
    [hashtable]$return = @{ }

    ConnectToDB -servername $au2matorDBServer -database $au2matorDBName 

    $Result = ExecuteSqlQuery -sqlquery "SELECT        RPM.Text AS Question, RP.Value
    FROM            dbo.Requests AS R INNER JOIN
                             dbo.RunbookParameterMappings AS RPM ON R.ServiceId = RPM.ServiceId INNER JOIN
                             dbo.RequestParameters AS RP ON RPM.ParameterName = RP.[Key] AND R.RequestId = RP.RequestId
    where RP.RequestId = '$RequestID' order by [Order]"
    
    $html = "<table><tr><td><b>Question</b></td><td><b>Answer</b></td></tr>"
    $html = "<table>"
    foreach ($row in $Result) { 
        $row
        $html += "<tr><td><b>" + $row.Question + "</b></td><td>" + $row.Value + "</td></tr>" 
    }
    $html += "</table>"

    $f_RequestInfo = ExecuteSqlQuery -sqlquery "select InitiatedBy, TargetUserId,[ApprovedBy], [ApprovedTime], Comment from Requests where RequestId =  '$RequestID'"

    $Connection.Close()
    Remove-Variable -Scope Global -Name Connection

    $f_SamInitiatedBy = $f_RequestInfo.InitiatedBy.Split("\")[1]
    $f_UserInitiatedBy = Get-ADUser -Identity $f_SamInitiatedBy -Properties Mail

    $f_SamTarget = $f_RequestInfo.TargetUserId.Split("\")[1]
    $f_UserTarget = Get-ADUser -Identity $f_SamTarget -Properties Mail

    $return.InitiatedBy = $f_RequestInfo.InitiatedBy
    $return.MailInitiatedBy = $f_UserInitiatedBy.mail
    $return.MailTarget = $f_UserTarget.mail
    $return.TargetUserId = $f_RequestInfo.TargetUserId
    $return.ApprovedBy = $f_RequestInfo.ApprovedBy
    $return.ApprovedTime = $f_RequestInfo.ApprovedTime
    $return.Comment = $f_RequestInfo.Comment
    $return.HTML = $HTML

    return $return
}

Function Get-MailContent ($RequestID, $RequestTitle, $EndDate, $TargetUserId, $InitiatedBy, $Status, $PortalURL, $RequestedBy, $AdditionalHTML, $InputHTML) {

    $f_RequestID = $RequestID
    $f_InitiatedBy = $InitiatedBy
    
    $f_RequestTitle = $RequestTitle
    $f_EndDate = $EndDate
    $f_RequestStatus = $Status
    $f_RequestLink = "$PortalURL/requeststatus?id=$RequestID"
    $f_RequestedBy = $RequestedBy
    $f_HTMLINFO = $AdditionalHTML
    $f_InputHTML = $InputHTML
    
    $f_SamInitiatedBy = $f_InitiatedBy.Split("\")[1]
    $f_UserInitiatedBy = Get-ADUser -Identity $f_SamInitiatedBy -Properties DisplayName
    $f_DisplaynameInitiatedBy = $f_UserInitiatedBy.DisplayName

    
    $HTML = @'    
    <table class="MsoNormalTable" style="width: 100.0%; mso-cellspacing: 1.5pt; background: #F7F8F3; mso-yfti-tbllook: 1184;" border="0" width="100%" cellpadding="0">
    <tbody>
    <tr style="mso-yfti-irow: 0; mso-yfti-firstrow: yes; mso-yfti-lastrow: yes;">
    <td style="padding: .75pt .75pt .75pt .75pt;" valign="top">&nbsp;</td>
    <td style="width: 450.0pt; padding: .75pt .75pt .75pt .75pt; box-sizing: border-box;" valign="top" width="600">
    <div style="box-sizing: border-box;">
    <table class="MsoNormalTable" style="width: 100.0%; mso-cellspacing: 0cm; background: white; border: solid #E9E9E9 1.0pt; mso-border-alt: solid #E9E9E9 .75pt; mso-yfti-tbllook: 1184; mso-padding-alt: 0cm 0cm 0cm 0cm;" border="1" width="100%" cellspacing="0" cellpadding="0">
    <tbody>
    <tr style="mso-yfti-irow: 0; mso-yfti-firstrow: yes;">
    <td style="border: none; background: #6ddc36; padding: 15.0pt 0cm 15.0pt 15.0pt;" valign="top">
    <p class="MsoNormal" style="line-height: 19.2pt;"><img src="https://au2mator.com/wp-content/uploads/2018/02/HPLogoau2mator-1.png" alt="" width="198" height="43" /></p>
    </td>
    </tr>
    <tr style="mso-yfti-irow: 1; box-sizing: border-box;">
    <td style="border: none; padding: 15.0pt 15.0pt 15.0pt 15.0pt; box-sizing: border-box;" valign="top">
    <table class="MsoNormalTable" style="width: 100.0%; mso-cellspacing: 0cm; mso-yfti-tbllook: 1184; mso-padding-alt: 0cm 0cm 0cm 0cm; box-sizing: border-box;" border="0" width="100%" cellspacing="0" cellpadding="0">
    <tbody>
    <tr style="mso-yfti-irow: 0; mso-yfti-firstrow: yes; mso-yfti-lastrow: yes; box-sizing: border-box;">
    <td style="padding: 0cm 0cm 15.0pt 0cm; box-sizing: border-box;" valign="top">
    <table class="MsoNormalTable" style="width: 100.0%; mso-cellspacing: 0cm; mso-yfti-tbllook: 1184; mso-padding-alt: 0cm 0cm 0cm 0cm; box-sizing: border-box;" border="0" width="100%" cellspacing="0" cellpadding="0">
    <tbody>
    <tr style="mso-yfti-irow: 0; mso-yfti-firstrow: yes; mso-yfti-lastrow: yes; box-sizing: border-box;">
    <td style="width: 55.0%; padding: 0cm 0cm 0cm 0cm; box-sizing: border-box;" valign="top" width="55%">
    <table class="MsoNormalTable" style="width: 100.0%; mso-cellspacing: 0cm; mso-yfti-tbllook: 1184; mso-padding-alt: 0cm 0cm 0cm 0cm;" border="0" width="100%" cellspacing="0" cellpadding="0">
    <tbody>
    <tr style="mso-yfti-irow: 0; mso-yfti-firstrow: yes;">
    <td style="width: 18.75pt; border-top: solid #E3E3E3 1.0pt; border-left: solid #E3E3E3 1.0pt; border-bottom: none; border-right: none; mso-border-top-alt: solid #E3E3E3 .75pt; mso-border-left-alt: solid #E3E3E3 .75pt; padding: 0cm 0cm 0cm 0cm; box-sizing: border-box;" width="25">
    <p class="MsoNormal" style="text-align: center; line-height: 19.2pt;" align="center">&nbsp;</p>
    </td>
    <td style="border-top: solid #E3E3E3 1.0pt; border-left: none; border-bottom: none; border-right: solid #E3E3E3 1.0pt; mso-border-top-alt: solid #E3E3E3 .75pt; mso-border-right-alt: solid #E3E3E3 .75pt; padding: 0cm 0cm 3.75pt 0cm; font-color: #0000;"><strong>End Date</strong>: ##EndDate</td>
    </tr>
    <tr style="mso-yfti-irow: 1;">
    <td style="border-top: solid #E3E3E3 1.0pt; border-left: solid #E3E3E3 1.0pt; border-bottom: none; border-right: none; mso-border-top-alt: solid #E3E3E3 .75pt; mso-border-left-alt: solid #E3E3E3 .75pt; padding: 0cm 0cm 0cm 0cm;">
    <p class="MsoNormal" style="text-align: center; line-height: 19.2pt;" align="center">&nbsp;</p>
    </td>
    <td style="border-top: solid #E3E3E3 1.0pt; border-left: none; border-bottom: none; border-right: solid #E3E3E3 1.0pt; mso-border-top-alt: solid #E3E3E3 .75pt; mso-border-right-alt: solid #E3E3E3 .75pt; padding: 0cm 0cm 3.75pt 0cm;"><strong>Status</strong>: ##Status</td>
    </tr>
    <tr style="mso-yfti-irow: 2; mso-yfti-lastrow: yes;">
    <td style="border: solid #E3E3E3 1.0pt; border-right: none; mso-border-top-alt: solid #E3E3E3 .75pt; mso-border-left-alt: solid #E3E3E3 .75pt; mso-border-bottom-alt: solid #E3E3E3 .75pt; padding: 0cm 0cm 3.75pt 0cm;">
    <p class="MsoNormal" style="text-align: center; line-height: 19.2pt;" align="center">&nbsp;</p>
    </td>
    <td style="border: solid #E3E3E3 1.0pt; border-left: none; mso-border-top-alt: solid #E3E3E3 .75pt; mso-border-bottom-alt: solid #E3E3E3 .75pt; mso-border-right-alt: solid #E3E3E3 .75pt; padding: 0cm 0cm 3.75pt 0cm;"><strong>Requested By</strong>: ##RequestedBy</td>
    </tr>
    </tbody>
    </table>
    </td>
    <td style="width: 5.0%; padding: 0cm 0cm 0cm 0cm; box-sizing: border-box;" width="5%">
    <p class="MsoNormal" style="line-height: 19.2pt;"><span style="font-size: 9.0pt; font-family: 'Helvetica',sans-serif; mso-fareast-font-family: 'Times New Roman';">&nbsp;</span></p>
    </td>
    <td style="width: 40.0%; padding: 0cm 0cm 0cm 0cm; box-sizing: border-box;" valign="top" width="40%">
    <table class="MsoNormalTable" style="width: 100.0%; mso-cellspacing: 0cm; background: #FAFAFA; mso-yfti-tbllook: 1184; mso-padding-alt: 0cm 0cm 0cm 0cm;" border="0" width="100%" cellspacing="0" cellpadding="0">
    <tbody>
    <tr style="mso-yfti-irow: 0; mso-yfti-firstrow: yes; mso-yfti-lastrow: yes;">
    <td style="width: 100.0%; border: solid #E3E3E3 1.0pt; mso-border-alt: solid #E3E3E3 .75pt; padding: 7.5pt 0cm 1.5pt 3.75pt;" width="100%">
    <p style="text-align: center;" align="center"><span style="font-size: 10.5pt; color: #959595;">Request ID</span></p>
    <p class="MsoNormal" style="text-align: center;" align="center">&nbsp;</p>
    <p style="text-align: center;" align="center"><u><span style="font-size: 12.0pt; color: black;"><a href="##RequestLink"><span style="color: black;">##REQUESTID</span></a></span></u></p>
    <p class="MsoNormal" style="text-align: center;" align="center"><span style="mso-fareast-font-family: 'Times New Roman';">&nbsp;</span></p>
    </td>
    </tr>
    </tbody>
    </table>
    </td>
    </tr>
    </tbody>
    </table>
    </td>
    </tr>
    </tbody>
    </table>
    </td>
    </tr>
    <tr style="mso-yfti-irow: 2; box-sizing: border-box;">
    <td style="border: none; padding: 0cm 15.0pt 15.0pt 15.0pt; box-sizing: border-box;" valign="top">
    <table class="MsoNormalTable" style="width: 100.0%; mso-cellspacing: 0cm; mso-yfti-tbllook: 1184; mso-padding-alt: 0cm 0cm 0cm 0cm; box-sizing: border-box;" border="0" width="100%" cellspacing="0" cellpadding="0">
    <tbody>
    <tr style="mso-yfti-irow: 0; mso-yfti-firstrow: yes; box-sizing: border-box;">
    <td style="padding: 0cm 0cm 15.0pt 0cm; box-sizing: border-box;" valign="top">
    <p class="MsoNormal" style="line-height: 19.2pt;"><strong><span style="font-size: 10.5pt; font-family: 'Helvetica',sans-serif; mso-fareast-font-family: 'Times New Roman';">Dear ##UserDisplayname,</span></strong></p>
    </td>
    </tr>
    <tr style="mso-yfti-irow: 1; box-sizing: border-box;">
    <td style="padding: 0cm 0cm 15.0pt 0cm; box-sizing: border-box;" valign="top">
    <p class="MsoNormal" style="line-height: 19.2pt;"><span style="font-size: 10.5pt; font-family: 'Helvetica',sans-serif; mso-fareast-font-family: 'Times New Roman';">the Request <strong>"##RequestTitle"</strong> has been finished.<br /> <br /> Please see the description for detailed information.<br /><b>##HTMLINFO&nbsp;</b><br /></span></p>
    <div>&nbsp;</div>
    <div>See the Details of the Request</div>
    <div>##InputHTML</div>
    <div>&nbsp;</div>
    <div>&nbsp;</div>
    Kind regards,<br /> au2mator Self Service Team
    <p>&nbsp;</p>
    </td>
    </tr>
    <tr style="mso-yfti-irow: 2; mso-yfti-lastrow: yes; box-sizing: border-box;">
    <td style="padding: 0cm 0cm 15.0pt 0cm; box-sizing: border-box;" valign="top">
    <p class="MsoNormal" style="text-align: center; line-height: 19.2pt;" align="center"><span style="font-size: 10.5pt; font-family: 'Helvetica',sans-serif; mso-fareast-font-family: 'Times New Roman';"><a style="border-radius: 3px; -webkit-border-radius: 3px; -moz-border-radius: 3px; display: inline-block;" href="##RequestLink"><strong><span style="color: white; border: solid #50D691 6.0pt; padding: 0cm; background: #50D691; text-decoration: none; text-underline: none;">View your Request</span></strong></a></span></p>
    </td>
    </tr>
    </tbody>
    </table>
    </td>
    </tr>
    <tr style="mso-yfti-irow: 3; mso-yfti-lastrow: yes; box-sizing: border-box;">
    <td style="border: none; padding: 0cm 0cm 0cm 0cm; box-sizing: border-box;" valign="top">
    <table class="MsoNormalTable" style="width: 100.0%; mso-cellspacing: 0cm; background: #333333; mso-yfti-tbllook: 1184; mso-padding-alt: 0cm 0cm 0cm 0cm;" border="0" width="100%" cellspacing="0" cellpadding="0">
    <tbody>
    <tr style="mso-yfti-irow: 0; mso-yfti-firstrow: yes; mso-yfti-lastrow: yes; box-sizing: border-box;">
    <td style="width: 50.0%; border: none; border-right: solid lightgrey 1.0pt; mso-border-right-alt: solid lightgrey .75pt; padding: 22.5pt 15.0pt 22.5pt 15.0pt; box-sizing: border-box;" valign="top" width="50%">&nbsp;</td>
    <td style="width: 50.0%; padding: 22.5pt 15.0pt 22.5pt 15.0pt; box-sizing: border-box;" valign="top" width="50%">&nbsp;</td>
    </tr>
    </tbody>
    </table>
    </td>
    </tr>
    </tbody>
    </table>
    </div>
    </td>
    <td style="padding: .75pt .75pt .75pt .75pt; box-sizing: border-box;" valign="top">&nbsp;</td>
    </tr>
    </tbody>
    </table>
    <p class="MsoNormal"><span style="mso-fareast-font-family: 'Times New Roman';">&nbsp;</span></p>
'@

    $html = $html.replace('##REQUESTID', $f_RequestID).replace('##UserDisplayname', $f_DisplaynameInitiatedBy).replace('##RequestTitle', $f_RequestTitle).replace('##EndDate', $f_EndDate).replace('##Status', $f_RequestStatus).replace('##RequestedBy', $f_RequestedBy).replace('##HTMLINFO', $f_HTMLINFO).replace('##InputHTML', $f_InputHTML).replace('##RequestLink', $f_RequestLink)

    return $html
}

Function Send-ServiceMail ($HTMLBody, $ServiceName, $Recipient, $RequestID, $RequestStatus) {
    $f_Subject = "au2mator - $ServiceName Request [$RequestID] - $RequestStatus"

    if ($SMPTAuthentication) {
        $f_secpasswd = ConvertTo-SecureString $SMTPPassword -AsPlainText -Force
        $f_mycreds = New-Object System.Management.Automation.PSCredential ($SMTPUser, $f_secpasswd)
    
        if ($EnableSSLforSMTP) {
            Send-MailMessage -SmtpServer $SMTPServer -To $Recipient -From $SMTPSender -Subject $f_Subject -Body $HTMLBody -BodyAsHtml -Priority high -Credential $f_mycreds -UseSsl
        }
        else {
            Send-MailMessage -SmtpServer $SMTPServer -To $Recipient -From $SMTPSender -Subject $f_Subject -Body $HTMLBody -BodyAsHtml -Priority high -Credential $f_mycreds
        }
    }
    else {
        if ($EnableSSLforSMTP) {
            Send-MailMessage -SmtpServer $SMTPServer -To $Recipient -From $SMTPSender -Subject $f_Subject -Body $HTMLBody -BodyAsHtml -Priority high -UseSsl
        }
        else {
            Send-MailMessage -SmtpServer $SMTPServer -To $Recipient -From $SMTPSender -Subject $f_Subject -Body $HTMLBody -BodyAsHtml -Priority high
        } 
    }   
}

##Custom Functions

Function Send-ExportMail ($HTMLBody, $Subject, $Recipient, $RequestStatus, $File) {
    $f_Subject = $Subject

    if ($SMPTAuthentication) {
        $f_secpasswd = ConvertTo-SecureString $SMTPPassword -AsPlainText -Force
        $f_mycreds = New-Object System.Management.Automation.PSCredential ($SMTPUser, $f_secpasswd)
    
        if ($EnableSSLforSMTP) {
            Send-MailMessage -SmtpServer $SMTPServer -To $Recipient -From $SMTPSender -Subject $f_Subject -Body $HTMLBody -BodyAsHtml -Priority high -Credential $f_mycreds -UseSsl -Attachments $File
        }
        else {
            Send-MailMessage -SmtpServer $SMTPServer -To $Recipient -From $SMTPSender -Subject $f_Subject -Body $HTMLBody -BodyAsHtml -Priority high -Credential $f_mycreds -Attachments $File
        }
    }
    else {
        if ($EnableSSLforSMTP) {
            Send-MailMessage -SmtpServer $SMTPServer -To $Recipient -From $SMTPSender -Subject $f_Subject -Body $HTMLBody -BodyAsHtml -Priority high -UseSsl -Attachments $File
        }
        else {
            Send-MailMessage -SmtpServer $SMTPServer -To $Recipient -From $SMTPSender -Subject $f_Subject -Body $HTMLBody -BodyAsHtml -Priority high -Attachments $File
        } 
    }   
}


##################################################################################################### 
Function Cleanup() { 
    #$ErrorActionPreference = "SilentlyContinue"    #Depending on when this is called, some variables may not be initialized and clearing could throw benign error. Supress.  
    Write-au2matorLog -Type INFO -Text "Performing cleanup..."
    #Cleanup 
    Get-Variable | Where-Object { $StartupVariables -notcontains $_.Name } | Foreach-Object { Remove-Variable -Name "$($_.Name)" -Force -Scope 1 } 
} 

########################################################################################################     
#   Will clean up names/strings with special characters (like URLs and Network paths) 
Function CleanName { 
    Param( 
        [string]$uglyString 
    ) 
    # Remove problematic characters and leading/trailing spaces 
    $prettyString = (($uglyString.Replace(':', '_')).Replace('/', '_')).Replace('\', '_').Trim() 

    # If the string has been modified, output that info 
    If ($uglyString -ne $prettyString) { 
        Write-au2matorLog -Type INFO -Text  "There was a small problem with the characters in this parameter: [$($uglyString)]..." 
        Write-au2matorLog -Type INFO -Text  "Original Name:`t`t$($uglyString)" 
        Write-au2matorLog -Type INFO -Text  "Modified Name:`t`t$($prettyString)"  
    } 

    Return $prettyString 
    #> 
} 
########################################################################################################     
#    Function MergeFiles 
#    Will find .csv files and merge them together.  
Function MergeFiles { 
    Param( 
        [string]$strPath, 
        [string]$strOutputFileName 
    ) 


    $strOutputFilePath = (Join-Path $strPath $strOutputFileName) 
    # If output file already exists, remove it.  
    If (Test-Path $strOutputFilePath -PathType Leaf) {  
        Write-au2matorLog -Type INFO -Text  "Output file [ $($strOutputFilePath) ] already exists. Removing..." 
        Remove-Item -Path $strOutputFilePath -Force  
    } 

    If (Test-Path $strOutputFilePath) {  
        Write-au2matorLog -Type ERROR -Text  "Cannot remove $strOutputFilePath and therefore cannot generate merged output file."
        Write-au2matorLog -Type ERROR -Text  "Remove this file first: [ $($strOutputFilePath) ]" 
        Write-au2matorLog -Type ERROR -Text  "Exiting ! " 
        Cleanup 
        Exit 
    } 

    Get-ChildItem -Path $strPath -File -Filter *.csv -Exclude $strOutputFileName -Recurse | ForEach-Object { 
        $intThisHeaderLength = (Get-Content -LiteralPath $_.FullName)[0].Length  
        If ($intThisHeaderLength -gt $intLongestHeaderLength) { 
            $objLongestHeaderFile = $_ 
            $intLongestHeaderLength = $intThisHeaderLength 
        } 
    } 

    Write-au2matorLog -Type INFO -Text  "Has largest set of file headers: [ $($objLongestHeaderFile.FullName) ] " 
    # Create the master merge file seeded by the data from the existing CSV file with the most headers out of the entire set of CSVs. 
    Try { 
        Get-Content $objLongestHeaderFile.FullName | Out-File -LiteralPath $strOutputFilePath -Force -Encoding UTF8 
    }
    Catch { 
        Write-au2matorLog -Type ERROR -Text  $error[0] 
        Write-au2matorLog -Type INFO -Text  "Something is wrong with this path [$($strOutputFilePath)]."
        Write-au2matorLog -Type INFO -Text  "Exiting..."  
        Exit 
    } 
    # Iterate through all of the CSVs, append all of them into the master (except for the one already used as the seed above and except for the master merge file itself.) 
    $i = 0 
    $tempArray = @() 
    Get-ChildItem -Path $strPath -File -Filter *.csv -Exclude "merged*" -Recurse | ForEach-Object {  
        If ( ( $_.FullName -eq $objLongestHeaderFile.FullName ) -or ($_.FullName -like (Get-Item $strOutputFilePath).FullName) ) { 
            Write-au2matorLog -Type INFO -Text  "Skip this file:"
            Write-au2matorLog -Type INFO -Text  "$($_.FullName)" 
        } 
        Else { 
            Write-au2matorLog -Type INFO -Text  "Merge this file: "
            Write-au2matorLog -Type INFO -Text  "$($_.FullName)"
            $tempArray += ((((Get-Content -Raw -Path $_.FullName) -Replace "\n", " " ) -Split "\r") | Select-Object -Skip 1 ) 
            $i++ 
        } 
    } 
    $tempArray | Out-File -LiteralPath $strOutputFilePath -Append -Encoding UTF8 
    "" # Cheap formatting 
    Write-au2matorLog -Type INFO -Text  "Total files merged: "
    Write-au2matorLog -Type INFO -Text  "$i"
    Write-au2matorLog -Type INFO -Text  "Master output file: "
    Write-au2matorLog -Type INFO -Text  "$strOutputFilePath"
} # EndFunction 
################################################################################################### 

Function MakeObject { 
    Param( 
        [string]$strMergedFilePath 
    ) 
    $mainArray = @() 
    [string]$rootFolder = Split-Path $strMergedFilePath -Parent 

    $tmpFileName = "ExportEffectiveMonitoringConfiguration.ps1_DisplayNamesCSV.tmp" 
    Try { 
        [string]$savedDNs = (Join-Path $env:Temp $tmpFileName )         
        New-Item -ItemType File -Path $savedDNs -ErrorAction SilentlyContinue 
    }
    Catch { 
        [string]$savedDNs = (Join-Path $rootFolder $tmpFileName ) 
    } 
    If ($ClearCache -eq "True") { 
        Write-au2matorLog -Type INFO -Text  "Removing saved DisplayNames file: [$($savedDNs)]"
        Remove-Item -Path $savedDNs -Force 
    } 

    If (!($strMergedFilePath)) { 
        Write-au2matorLog -Type INFO -Text  "Cannot find [ $($strMergedFilePath) ] and therefore cannot compile object for Grid View."
        Write-au2matorLog -Type INFO -Text  "Exiting..." 
        Cleanup 
        Exit 
    } 

    $Headers = @() 
    $Headers = (Get-Content -LiteralPath $strMergedFilePath | Select-Object -First 1).Split('|') 
    $FileContents = (Get-Content -LiteralPath $strMergedFilePath | Select-Object -Skip 1 ).Replace("`0", '') 
    $r = 1 

    <# The Export-SCOMEffectiveMonitoringConfiguration cmdlet does not include DisplayName in it's default output set. 
Querying the SDK for the workflow DisplayName is expensive.  In the code below we try to benefit from a saved  
list of Name->DisplayName pairs. If the list does not already exist, we will create one. If it does already  
exist, we will import it into a hash table for fast DisplayName lookup while building the rows of the master file.  
#> 
    $DNHash = @{ } 
    Try { 
        [System.Object[]]$arrDN = (Import-Csv -Path $savedDNs -ErrorAction SilentlyContinue) 
    }
    Catch { 
        $arrDN = @() 
    } 
    # If a previous list of Name/DisplayName pairs exists, let's use it to build our fast hash table. 
    If ([bool]$arrDN ) { 
        ForEach ($item in $arrDN) { 
            $DNHash.Add($item.'Rule/Monitor Name', $item.'Rule/Monitor DisplayName') 
        } 
    } 
    $arrTmpDN = @() 
    ForEach ($Row in $FileContents) { 
        $percent = [math]::Round(($r / $FileContents.count * 100), 0)  
        #Write-Progress -Activity "** What's happening? **" -status "Formatting your data! [Percent: $($percent)]" -percentComplete $percent 
        If ($Row.Length -le 1) { Continue; } 
        $c = 0 
        $arrRow = @() 
        $arrRow = $Row.Split('|') 

        # If the ForEach has already executed one iteration and thus the full object template has already been created,  
        # duplicate the template instead of building a new object and adding members to it for each column. This is about 3x faster than building the object every iteration. 
        If ([bool]($templateObject)) { 
            $object = $templateObject.PsObject.Copy() 
            $object.Index = $r.ToString("0000") 
        } 
        Else { 
            $object = New-Object -TypeName PSObject 
            $object | Add-Member -MemberType NoteProperty -Name "Index" -Value $r.ToString("0000") 
        }         
        ForEach ($Column in $Headers) { 
            If ( ($arrRow[$c] -eq '') -or ($arrRow[$c] -eq ' ') ) { $arrRow[$c] = 'N/A' } 
            # Some header values repeat. If header already exists, give it a unique name 
            [int]$Position = 1 
            $tempColumn = $Column 
            # The first 10 columns are unique. However, beyond 10, the column names repeat: 
            #  Parameter Name, Default Value, Effective Value 
            # A clever way to assign each set of repeats a unique name is to append an incremental instance number.  
            # Each set (of 3 column names) gets an occurance number provided by the clever math below. 
            # Example: Parameter Name1, Default Value1, Effective Value1, Parameter Name2, Default Value2, Effective Value2 
            If ($c -ge 10) { 
                $Position = [System.Math]::Ceiling(($c / 3) - 3) 
                $tempColumn = $Column + "$Position" 
            } 
            If ([bool]($templateObject)) { 
                $object.$tempColumn = $arrRow[$c] 
            } 
            Else { $object | Add-Member -MemberType NoteProperty -Name $tempColumn -Value "$($arrRow[$c])" } 

            If ($Column -eq 'Rule/Monitor Name') { 
                # If DisplayName (DN) does not already exist in set 
                If (-not [bool]($DN = $DNHash.($arrRow[$c])) ) { 
                    # Find the DisplayName 
                    switch ($arrRow[7]) { #Assuming this column header is consistently "Type"  
                        'Monitor' { 
                            $DN = (Get-SCOMMonitor -Name $arrRow[$c]).DisplayName  
                        } 
                        'Rule' { 
                            $DN = (Get-SCOMRule -Name $arrRow[$c]).DisplayName 
                        } 
                        Default { Write-Host "SWITCH DEFAULT IN FUNCTION: 'MAKEOBJECT', SOMETHING IS WRONG." -F Red -B Yellow } 
                    } 

                    # If no DN exists for the workflow, set a default 
                    If (-Not([bool]$DN)) { 
                        $DN = "N/A" 
                    } 
                    Else { 
                        $DNHash.Add($arrRow[$c], $DN)  
                    } 
                } 
                # DN Exists, add it to the hash table for fast lookup. Also add it to the catalog/array of known DNs so it can be saved and used again 
                #  next time for fast lookup. 

                If ([bool]($templateObject)) { 
                    $object.'Rule/Monitor DisplayName' = $DN 
                } 
                Else { $object | Add-Member -MemberType NoteProperty -Name "Rule/Monitor DisplayName" -Value $DN } 
            } 
            $c++ 
        } 
        $r++ 
        $mainArray += $object 
        If (-not [bool]($templateObject)) { 
            $templateObject = $object.PsObject.Copy() 
        } 
        Remove-Variable -name object, tmpObj, DN -ErrorAction SilentlyContinue 
    } 
    # Build a simple array to hold unique Name,DisplayName values so that it can be exported easily to a CSV file.  
    # This cached csv file will significantly speed up the script next time it runs. 
    ForEach ($Key in $DNHash.Keys) { 
        $tmpObj = New-Object -TypeName PSObject 
        $tmpObj | Add-Member -MemberType NoteProperty -Name "Rule/Monitor Name" -Value $Key 
        $tmpObj | Add-Member -MemberType NoteProperty -Name "Rule/Monitor DisplayName" -Value $DNHash.$Key 
        $arrTmpDN += $tmpObj 
    } 
    $mainArray | Export-Csv -Path $strMergedFilePath -Force -Encoding UTF8 -Delimiter '|' -NoTypeInformation 
    $arrTmpDN | Export-Csv -Path $savedDNs -Force -Encoding UTF8 -NoTypeInformation 
    Return $mainArray 
} 
################################################################################################### 

#endregion Functions


#region Script
Write-au2matorLog -Type INFO -Text "Start Script"
Write-au2matorLog -Type INFO -Text "Write Variables and Values"
Get-Params

if ($DoImportPSSession) {

    Write-au2matorLog -Type INFO -Text "Import-Pssession"
    $PSSession = New-PSSession -ComputerName $OpsmgrServer
    Import-PSSession -Session $PSSession -DisableNameChecking -ArgumentList '.\SCOM - Start Maintenance Mode' -AllowClobber
}
else {
        
}

Write-au2matorLog -Type INFO -Text "Import SCOM PS Module"
Import-Module OperationsManager



If (!(Test-Path $TargetFolder)) {  
    Write-au2matorLog -Type INFO -Text  "TargetFolder [ $($TargetFolder) ] does not exist. Creating it now..." 
    new-item -ItemType Directory -Path $TargetFolder 
    If (!(Test-Path $TargetFolder)) { 
        Write-au2matorLog -Type ERROR -Text  "Unable to create TargetFolder: $TargetFolder. Exiting." 
        Cleanup 
        Exit 
    } 
    Else {  
        Write-au2matorLog -Type INFO -Text  "Created TargetFolder successfully. " 
    } 
} 


# If group name is provided... 
If ($c_SCOMGroupName) { 
    $choice = 'group' 
    $objects = @(Get-SCOMGroup -DisplayName $c_SCOMGroupName | Get-SCOMClassInstance)  
    If (-not($objects)) {
        Write-au2matorLog -Type ERROR -Text  "Unable to get group: [ $($c_SCOMGroupName) ]." 
        Write-au2matorLog -Type INFO -Text  "To troubleshoot, run this command:`n`n  Get-SCOMGroup -DisplayName '$c_SCOMGroupName' | Get-SCOMClassInstance `n"  
        Write-au2matorLog -Type ERROR -Text  "Exiting..."; 
        $ErrorCount = 1 
        Cleanup 
        Exit  
    } 
    Else { 
        Write-au2matorLog -Type INFO -Text  "Success getting group: [ $($c_SCOMGroupName) ]." 
    } 
    $TempName = $c_SCOMGroupName 
    $count = $objects.GetRelatedMonitoringObjects().Count 
    "" # Cheap formatting 
    "" 
    Write-au2matorLog -Type INFO -Text  "This will output ALL monitoring configuration for group: " 
    Write-au2matorLog -Type INFO -Text  "["
    Write-au2matorLog -Type INFO -Text  "$($c_SCOMGroupName)" 
    Write-au2matorLog -Type INFO -Text  "]" 

    Write-au2matorLog -Type INFO -Text  "There are: " 
    Write-au2matorLog -Type INFO -Text  "["
    Write-au2matorLog -Type INFO -Text  "$($count)" 
    Write-au2matorLog -Type INFO -Text  "]"
    Write-au2matorLog -Type INFO -Text  " nested objects in that group."
    Write-au2matorLog -Type INFO -Text  "This might take a little while depending on how large the group is and how many hosted objects exist!"
    "" # Cheap formatting 
} 

# If ID is provided... 
ElseIf ($c_ID) { 
    $choice = 'ID' 
    Write-au2matorLog -Type INFO -Text  "Getting class instance with ID: [ $($c_ID) ] " 
    $objects = (Get-SCOMClassInstance -Id $c_ID) 
    If (-not($objects)) { 
        Write-au2matorLog -Type ERROR -Text  "Unable to get class instance for ID: [ $($c_ID) ] " 
        Write-au2matorLog -Type INFO -Text  "To troubleshoot, use this command:`n`n  Get-SCOMClassInstance -Id '$c_ID' `n" 
        Write-au2matorLog -Type ERROR -Text  "Exiting...";
        $ErrorCount = 1  
        Cleanup 
        Exit  
    } 
    Else { 
        Write-au2matorLog -Type INFO -Text "Success getting class instance with ID: [ $($c_ID) ], DisplayName: [ $($c_ID.DisplayName) ]." 
    } 
    $TempName = $c_ID 
    $count = $objects.GetRelatedMonitoringObjects().Count 
    "" # Cheap formatting 
    "" 
    Write-au2matorLog -Type INFO -Text  "This will output ALL monitoring configuration for object: "
    Write-au2matorLog -Type INFO -Text  "["
    Write-au2matorLog -Type INFO -Text  "$($objects.DisplayName) , "
    Write-au2matorLog -Type INFO -Text  "ID: $c_ID "
    Write-au2matorLog -Type INFO -Text  "]"
    Write-au2matorLog -Type INFO -Text  "There are: "
    Write-au2matorLog -Type INFO -Text  "["
    Write-au2matorLog -Type INFO -Text  "$($count)" 
    Write-au2matorLog -Type INFO -Text  "]"
    Write-au2matorLog -Type INFO -Text  " related monitoring objects."
    Write-au2matorLog -Type INFO -Text  "This might take a little while depending on how hosted objects exist !"
    "" # Cheap formatting 
} 
# Assume individul computer name is provided... 
ElseIf ($c_ComputerName) {  
    $choice = 'ComputerName' 
    # $objects = @(Get-SCOMClass -Name "Microsoft.Windows.Computer" | Get-SCOMClassInstance | Where-Object {$c_ComputerName -contains $_.DisplayName } )  
    # This approach should prove to be more efficient for environments with more than 40-ish Computers/agents. 
    $ClassName = 'Microsoft.Windows.Computer' 
    $ComputerClass = (Get-SCOMClass -Name $ClassName) 
    If (-not($ComputerClass)) { 
        Write-au2matorLog -Type ERROR -Text  "Unable to get class: [ $ClassName ]."  
        Write-au2matorLog -Type INFO -Text  "To troubleshoot, use this command:`n`n  Get-SCOMClass -Name '$ClassName' `n" 
        Write-au2matorLog -Type ERROR -Text "Exiting...";  
        $ErrorCount = 1
        Cleanup 
        Exit  
    } 
    Else { 
        Write-au2matorLog -Type INFO -Text  "Success getting class object with name: [ $($ClassName) ]." 
    } 

    Write-au2matorLog -Type INFO -Text  "Getting class instance of [ $($ClassName) ] with DisplayName of [ $($c_ComputerName) ]..." 
    $objects = @(Get-SCOMClassInstance -DisplayName $c_ComputerName | Where-Object { $_.LeastDerivedNonAbstractManagementPackClassId -like $ComputerClass.Id.Guid } ) 
    If (-not($objects)) { 
        Write-au2matorLog -Type ERROR -Text  "Unable to get class instance for DisplayName: [ $($c_ComputerName) ] " 
        Write-au2matorLog -Type INFO -Text  "To troubleshoot, use this command:`n`n   `$ComputerClass = (Get-SCOMClass -Name '$ClassName') " 
        Write-au2matorLog -Type INFO -Text  "   Get-SCOMClassInstance -DisplayName '$c_ComputerName' | Where-Object {`$_.LeastDerivedNonAbstractManagementPackClassId -like `$ComputerClass.Id.Guid} `n" 
        Write-au2matorLog -Type ERROR -Text  "Exiting...";  
        $ErrorCount = 1
        Cleanup 
        Exit  
    } 
    Else { 
        Write-au2matorLog -Type INFO -Text  "Success getting class instance for DisplayName: [ $($c_ComputerName) ] " 
    } 
    $TempName = $c_ComputerName 
    $count = $objects.GetRelatedMonitoringObjects().Count 

    "" # Cheap formatting 
    "" 
    Write-au2matorLog -Type INFO -Text  "This will output ALL monitoring configuration for Computer: "
    Write-au2matorLog -Type INFO -Text  "["
    Write-au2matorLog -Type INFO -Text  "$($objects.DisplayName)"
    Write-au2matorLog -Type INFO -Text  "]"
    Write-au2matorLog -Type INFO -Text  "There are: "
    Write-au2matorLog -Type INFO -Text  "["
    Write-au2matorLog -Type INFO -Text  "$($count)"
    Write-au2matorLog -Type INFO -Text  "]"
    Write-au2matorLog -Type INFO -Text  " related monitoring objects."
    Write-au2matorLog -Type INFO -Text  "This might take a little while depending on how many hosted objects exist !"
    "" # Cheap formatting 
} 
Else { 
    #This should never happen because of parameter validation 
    Write-au2matorLog -Type INFO -Text  "No input provided. Exiting..." 
    Cleanup 
    Exit 
} 


# If no OutputFileName exists, then simply use the DisplayName of the class instance. 
If (-not($OutputFileName)) { 
    $OutputFileName = "Merged_" + $TempName + ".csv" 
} 
Else { 
    $tempIndex = $OutputFileName.LastIndexOf('.') 
    If ($tempIndex -gt 1) { $temp = $OutputFileName.Substring(0, $tempIndex) } 
    Else { $temp = $OutputFileName } 
    $OutputFileName = "Merged_" + $temp + ".csv" 
} 


# Iterators used for nicely formatted output. 

# Iterate through the objects (including hosted instances) and dig out all related configs for rules/monitors. 
#$objects | % 
Foreach ($O in $objects)
{
    $DN = (CleanName -uglyString $O.DisplayName) 
    $path = (Join-Path $TargetFolder "($( CleanName -uglyString $O.Path ))_$($DN).csv" ) 
    Export-SCOMEffectiveMonitoringConfiguration -Instance $O -Path $path 
    Write-au2matorLog -Type INFO -Text  "$($i): " 
    Write-au2matorLog -Type INFO -Text  "["
    Write-au2matorLog -Type INFO -Text  "$($O.Path)"
    Write-au2matorLog -Type INFO -Text  "]"
    Write-au2matorLog -Type INFO -Text  " $($O.FullName)"
    $r = 1   #for progress bar calculation below 

    $related = @($O.GetRelatedMonitoringObjects()) 
    Write-au2matorLog -Type INFO -Text  "There are $($related.Count) 'related' monitoring objects for $($O.DisplayName)." 
    #$related | foreach ` 
    foreach ($r in $related)
    { 
        #$percent = [math]::Round((($r / $related.Count) * 100), 0) 
        #Write-Progress -Activity "** What's happening? **" -status "Getting your data. Be patient! [Percent: $($percent)]" -PercentComplete $percent 
        $DN = (($($r.DisplayName).Replace(':', '_')).Replace('/', '_')).Replace('\', '_') 
        $path = (Join-Path $TargetFolder "($($r.Path))_$($DN).csv" ) 
        Export-SCOMEffectiveMonitoringConfiguration -Instance $r -Path $path 
        #Write-Host "$($i): " -ForegroundColor Cyan -NoNewline; ` 
        #Write-Host "[" -ForegroundColor Red -NoNewline; ` 
        #Write-Host "$($_.Path)" -ForegroundColor Yellow -NoNewline; ` 
        #Write-Host "]" -ForegroundColor Red -NoNewline; ` 
        #Write-Host " $($_.FullName)"  -ForegroundColor Green 
        #$i++   # formatting, total line numbers 
        #$r++   # this object's hosted items, for progress bar calculation above 
    } 

    #$o++   
} 

#    ------ Merge Operation  ------ 
MergeFiles -strPath $TargetFolder -strOutputFileName $OutputFileName 
#    ------ Merge Operation  ------ 

if ($c_OptionExport -eq "Mail")
{
    Write-au2matorLog -Type INFO -Text  "Export will be sent by Mail"

    Send-ExportMail -HTMLBody "See attached the export of $temp" -Subject "SCOM Configuration $temp" -Recipient $c_ExportMail -File ($TargetFolder+"\"+$OutputFileName)
}


if ($c_OptionExport -eq "File")
{
    Write-au2matorLog -Type INFO -Text  "Export will be stored at Drive"
    Move-Item -Path ($TargetFolder+"\"+$OutputFileName) -Destination $c_ExportPath
}

Write-au2matorLog -Type INFO -Text  "Cleanup Export Path"
Start-Sleep -Seconds 10
Get-ChildItem -Path $TargetFolder  | Remove-Item -Recurse -Force


if ($ErrorCount -eq 0) {
    $au2matorReturn = "SCOM Export finished"
    $AdditionalHTML = "<br>
        Export Type: $c_OptionExport
        <br>
        Export Dest: $c_ExportPath$c_ExportMail
        <br>
        "
    $Status = "COMPLETED"
}
else {
    $au2matorReturn = "SCOM Export failed, Error: $Error"
    $Status = "ERROR"
}

#endregion Script





#region Return
## return to au2mator Services

Write-au2matorLog -Type INFO -Text "Service finished"

if ($SendMailToInitiatedByUser) {    
    Write-au2matorLog -Type INFO -Text "Send Mail to Initiated By User"

    $UserInput = Get-UserInput -RequestID $RequestId
    $HTML = Get-MailContent -RequestID $RequestId -RequestTitle $Service -EndDate $UserInput.ApprovedTime -TargetUserId $TargetUserId -InitiatedBy $InitiatedBy -Status $Status -PortalURL $PortalURL -RequestedBy $InitiatedBy -AdditionalHTML $AdditionalHTML -InputHTML $UserInput.html
    Send-ServiceMail -HTMLBody $HTML -RequestID $RequestId -Recipient "$($UserInput.MailInitiatedBy)" -RequestStatus $Status -ServiceName $Service
}


if ($SendMailToTargetUser) {    
    Write-au2matorLog -Type INFO -Text "Send Mail to Target User"
    $UserInput = Get-UserInput -RequestID $RequestId
    $HTML = Get-MailContent -RequestID $RequestId -RequestTitle $Service -EndDate $UserInput.ApprovedTime -TargetUserId $TargetUserId -InitiatedBy $InitiatedBy -Status $Status -PortalURL $PortalURL -RequestedBy $InitiatedBy -AdditionalHTML $AdditionalHTML -InputHTML $UserInput.html
    Send-ServiceMail -HTMLBody $HTML -RequestID $RequestId -Recipient "$($UserInput.MailTarget)" -RequestStatus $Status -ServiceName $Service
}


return $au2matorReturn    
#endregion Return




