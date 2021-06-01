<# PDF2SMTP_PS

 Version 1.2 - 01.06.2021, AUTHOR: Florian Mann
 
.SYNOPSYS
 Takes data from redfile, converts it to pdf and sends a mail with pdf attached back to the user which printed the file

.NOTES

.SOURCES
    Based on PDF2SMTP 1.6 from Frank Carius, info@netatwork.de, 31. March 2004
    Remove-StringSpecialCharacter from Francois-Xavier Cat, https://github.com/lazywinadmin/PowerShell/tree/master/TOOL-Remove-StringSpecialCharacter
. Prerequisite
    Redmon installed to c:\Program Files\gs\redmon\
    Redmon printer configured
    Ghostscript installed in c:\Program Files\gs\version\   for example "C:\Program Files\gs\gs9.52\bin\gswin64c.exe"
    Active Directory
    AD Users mail attribut set (Find User Mail-Address according to its Active Directory samaccountname)

.TODO
 Notes:


 Ideas:


 Improvements:
  - Check why printing websites often causes errors. (At least Firefox printing dialog error, but pdf delivered by mail (sometimes only one page)) Edge seems to work well
  - ELSE { $Regex = "[^a-z A-Z\. \p{Nd}]+" } to cover more special chars
  - Log shrink

 Critical Failures:
    - none known

     
.CHANGELOG
    1.0
        - First Release
        - Tested with Ghostscript 9.52
    1.1
        - Added Failover Mailserver to use if the first one is on a blocklist or other errors
        - Check possible SPAM blocks
    1.2
        - Added Usernames to Errormails

#>

#Start time measurement
    $temptimestart = Get-Date

#Functions
    #Shrinks logfiles
        function LOGSHRINK
        {
              param (
                [parameter(Mandatory=$true)][string]$pdf2smtp_ps_log
              )
            try
            {
                if(Test-Path $pdf2smtp_ps_log)
                {
                    #Reuse logfile
                    $logfilecontent = Get-Content $pdf2smtp_ps_log | Select-Object -Last 500 #Ceep last 500 lines of the logfile
                    [IO.File]::WriteAllLines($pdf2smtp_ps_log, $logfilecontent)
                }        
            }
            catch
            { "PDF2SMTP_PS $(Get-Date) Logfile creation $($_.Exception.Message)">> $pdf2smtp_ps_log  }

        }
    
        #LOGSHRINK $pdf2smtp_ps_log #Problems with character presentation


    function Remove-StringSpecialCharacter {
            <#
        .SYNOPSIS
            This function will remove the special character from a string.
        .DESCRIPTION
            This function will remove the special character from a string.
            I'm using Unicode Regular Expressions with the following categories
            \p{L} : any kind of letter from any language.
            \p{Nd} : a digit zero through nine in any script except ideographic
            http://www.regular-expressions.info/unicode.html
            http://unicode.org/reports/tr18/
        .PARAMETER String
            Specifies the String on which the special character will be removed
        .PARAMETER SpecialCharacterToKeep
            Specifies the special character to keep in the output
        .EXAMPLE
            Remove-StringSpecialCharacter -String "^&*@wow*(&(*&@"
            wow
        .EXAMPLE
            Remove-StringSpecialCharacter -String "wow#@!`~)(\|?/}{-_=+*"
            wow
        .EXAMPLE
            Remove-StringSpecialCharacter -String "wow#@!`~)(\|?/}{-_=+*" -SpecialCharacterToKeep "*","_","-"
            wow-_*
        .NOTES
            Francois-Xavier Cat
            @lazywinadmin
            lazywinadmin.com
            github.com/lazywinadmin
        #>
            [CmdletBinding()]
            param
            (
                [Parameter(ValueFromPipeline)]
                [ValidateNotNullOrEmpty()]
                [Alias('Text')]
                [System.String[]]$String,

                [Alias("Keep")]
                #[ValidateNotNullOrEmpty()]
                [String[]]$SpecialCharacterToKeep
            )
            PROCESS {
                try {
                    IF ($PSBoundParameters["SpecialCharacterToKeep"]) {
                        $Regex = "[^\p{L}\p{Nd}"
                        Foreach ($Character in $SpecialCharacterToKeep) {
                            IF ($Character -eq "-") {
                                $Regex += "-"
                            }
                            else {
                                $Regex += [Regex]::Escape($Character)
                            }
                            #$Regex += "/$character"
                        }

                        $Regex += "]+"
                    } #IF($PSBoundParameters["SpecialCharacterToKeep"])
                    ELSE { $Regex = "[^\p{L}\p{Nd}]+" }

                    FOREACH ($Str in $string) {
                        Write-Verbose -Message "Original String: $Str"
                        $Str -replace $regex, ""
                    }
                }
                catch {
                    $PSCmdlet.ThrowTerminatingError($_)
                }
            } #PROCESS
    }

#DEBUG
    $debug = $true          #Write extra debug logs


#Handle special Chars and Space
        $env:REDMON_DOCNAME = Remove-StringSpecialCharacter -String $env:REDMON_DOCNAME -SpecialCharacterToKeep " ",".","_","-"
        #Check if REDMON_DOCNAME is empty after Remove-StringSpecialCharacter or longer then 150 Chars
        if($env:REDMON_DOCNAME -eq "" -or $env:REDMON_DOCNAME.Length -gt 150)
            {$env:REDMON_DOCNAME = "DOCUMENT_NAME" }
        
        if($debug)
            { "PDF2SMTP_PS $(Get-Date) DEBUG REDMON_DOCNAME after Handle special Chars: $($env:REDMON_DOCNAME)" >> $pdf2smtp_ps_log }

#Settings
        $sendtimemails = $true  #Send mail with time infos for long jobs
        $errormails = $true          #Send Mail on Error 
        
    #Mail
        $MailSmtpServer = "HOSTNAME-MAILSERVER"
        $MailFrom = "FROM-MAILADDRESS" # example: "john@doe.com"
        $ErrorMailsTo = "ERROR-TO-MAILADDRESS" # example: "john@doe.com"
        $sendtimemailsmailto = "TIMING-TO-MAILADDRESS" # Special Errormails for jobs with long durations. example: "john@doe.com"
        
        #SPAM-Check
            $blocklistServers = @(
                'zen.spamhaus.org'
                'pbl.spamhaus.org'
                'sbl.spamhaus.org'
                'spam.abuse.ch'
                'xbl.spamhaus.org'
            )
            #Your Sending IP for Mails
            $SendingMailSourceIP = "IP" # example: 111.222.111.222
                $reversedSendingMailIP = ($SendingMailSourceIP -split '\.')[3..0] -join '.'

        #Failover
        $MailSmtpServerFailover="HOSTNAME-FAILOVER-MAILSERVER"
        $MailuserFailover="FAILOVER-FROM-MAILADDRESS"
        $MailPwFailover="PASSWORD"

   #Log Path
        $pdf2smtp_ps_log = "c:\Program Files\gs\redmon\pdf2smtp_ps.log"

    #Error Log Path
        $pdf2smtp_ps_error_log = "c:\Program Files\gs\redmon\pdf2smtp_ps_error.log"

    #TEMP Path
        $temppath = "c:\Program Files\gs\redmon\temp"

    #Ghostscript Path
        $gsversion = "gs9.52"
        $gspath = "C:\Program Files\gs\$($gsversion)\bin\gswin64c.exe"
        if(!(Test-Path $gspath))
        {
            #Ghostscript not found
            if($errormails)
            { Send-MailMessage -SmtpServer $MailSmtpServer -UseSsl -Subject "PDF2SMTP ERROR: Ghostscript" -Body "ERROR: Ghostscript not found" -from $MailFrom -To $ErrorMailsTo -ErrorAction SilentlyContinue }
            "PDF2SMTP_PS $(Get-Date) ERROR: Ghostscript not found: $($gspath)">> $pdf2smtp_ps_log
            #LOG
                "PDF2SMTP_PS $(Get-Date) -----------------------------------">> $pdf2smtp_ps_log
                "PDF2SMTP_PS $(Get-Date) END PDF2SMTP_PS">> $pdf2smtp_ps_log
                "PDF2SMTP_PS $(Get-Date) -----------------------------------">> $pdf2smtp_ps_log
            exit
        }

    #Redfile Path
        $redfilepath = "C:\Program Files\gs\redmon\redfile.exe"
        if(!(Test-Path $redfilepath))
        {
            #Redfile not found
            if($errormails)
            { Send-MailMessage -SmtpServer $MailSmtpServer -UseSsl -Subject "PDF2SMTP ERROR: Redfile" -Body "ERROR: Redfile not found" -from $MailFrom -To $ErrorMailsTo -ErrorAction SilentlyContinue }
            "PDF2SMTP_PS $(Get-Date) ERROR: Redfile not found: $($redfilepath)">> $pdf2smtp_ps_log
            #LOG
                "PDF2SMTP_PS $(Get-Date) -----------------------------------">> $pdf2smtp_ps_log
                "PDF2SMTP_PS $(Get-Date) END PDF2SMTP_PS">> $pdf2smtp_ps_log
                "PDF2SMTP_PS $(Get-Date) -----------------------------------">> $pdf2smtp_ps_log

            exit
        }

    #PDF File
        $pdffile = "$($temppath)\$($env:REDMON_USER)-$($env:REDMON_DOCNAME)-$($env:REDMON_JOB).pdf"
        #PDF Settings
            switch($env:REDMON_PRINTER)
            {
                     #Printernames are used for PDF protection settings
                     "PDF2SMTP_CP" { $pdfprotect = "-4" } # all
                     "PDF2SMTP" { $pdfprotect = "-1852" } # view and print
            }
            default{ $pdfprotect = "-1852" }

            #Password for PDF protection
                $pdfpassword = "PASSWORD"

            # -3904 = view Only
            # -1856 = view and print
            # -4 = enable all
            # details http://casper.ghostscript.com/~ghostgum/pdftips.htm

    #PS File
        $psfile = "$($temppath)\job-$($env:REDMON_JOB).ps"


#Check if REDMON_* Variables are empty
if(!($env:REDMON_JOB -or $env:REDMON_PRINTER -or $env:REDMON_MACHINE -or $env:REDMON_USER -or $env:REDMON_DOCNAME))
{
    #Redmon Variables empty -> exit script
    "PDF2SMTP_PS $(Get-Date) ERROR: Empty mandatory REDMON_ variables">> $pdf2smtp_ps_log
    #LOG
        "PDF2SMTP_PS $(Get-Date) -----------------------------------">> $pdf2smtp_ps_log
        "PDF2SMTP_PS $(Get-Date) END PDF2SMTP_PS">> $pdf2smtp_ps_log
        "PDF2SMTP_PS $(Get-Date) -----------------------------------">> $pdf2smtp_ps_log
    exit
}


#LOG
    "PDF2SMTP_PS $(Get-Date) -----------------------------------">> $pdf2smtp_ps_log
    "PDF2SMTP_PS $(Get-Date) Begin PDF2SMTP_PS">> $pdf2smtp_ps_log
    "PDF2SMTP_PS $(Get-Date) -----------------------------------">> $pdf2smtp_ps_log

    if($debug)
    {
        "PDF2SMTP_PS $(Get-Date) DEBUG REDMON_JOB: $($env:REDMON_JOB)" >> $pdf2smtp_ps_log
        "PDF2SMTP_PS $(Get-Date) DEBUG REDMON_PORT: $($env:REDMON_PORT)" >> $pdf2smtp_ps_log
        "PDF2SMTP_PS $(Get-Date) DEBUG REDMON_PRINTER: $($env:REDMON_PRINTER)" >> $pdf2smtp_ps_log
        "PDF2SMTP_PS $(Get-Date) DEBUG REDMON_MACHINE: $($env:REDMON_MACHINE)" >> $pdf2smtp_ps_log
        "PDF2SMTP_PS $(Get-Date) DEBUG REDMON_USER: $($env:REDMON_USER)" >> $pdf2smtp_ps_log
        "PDF2SMTP_PS $(Get-Date) DEBUG REDMON_DOCNAME: $($env:REDMON_DOCNAME)" >> $pdf2smtp_ps_log
        "PDF2SMTP_PS $(Get-Date) DEBUG REDMON_SESSIONID: $($env:REDMON_SESSIONID)" >> $pdf2smtp_ps_log
    }

#LOG
    "PDF2SMTP_PS $(Get-Date) Start redirect redfile to $($psfile)">> $pdf2smtp_ps_log

#Start time measurement - redfile
    $tempredfiletimestart = Get-Date

#Redirect refile to .ps
    $allOutput = & $redfilepath $psfile # 2>&1
    #$stderr = $allOutput | ?{ $_ -is [System.Management.Automation.ErrorRecord] }

    if($debug)
    {
        #Log processing time - redfile
            $tempredfiletimeend = get-date
            $tempredfiletime = "$(($tempredfiletimeend - $tempredfiletimestart).Minutes):$(($tempredfiletimeend - $tempredfiletimestart).Seconds)"
            "PDF2SMTP_PS $(Get-Date) DEBUG Redfile processing time: $($tempredfiletime) Minute(s):Second(s)" >> $pdf2smtp_ps_log
    }
    if($debug)
        { "PDF2SMTP_PS $(Get-Date) DEBUG Redfile ExitCode $($allOutput.ExitCode)">>$pdf2smtp_ps_log } #Not Working ????

if(Test-Path $psfile)
{
    $psfilecreatonsuccess = $true
    #start "PDF2SMTP GSKonvert" /WAIT %pdf2smtpgs% -dNOPAUSE -dBATCH -dEncryptionR#3 -dKeyLength#128 -sDEVICE=pdfwrite -dPDFSETTINGS=/default -dPermissions#%pdfprotect% -sOwnerPassword=%pdfpassword% -dCompatibilityLevel=1.4 -sOutputFile=%PDFFILE% %PSFILE%

    #LOG
        "PDF2SMTP_PS $(Get-Date) Start Ghostscript">> $pdf2smtp_ps_log

    #Ghostscript Parameters
    $gssearchParams = @()
        $gssearchParams += "-dNOPAUSE"
        $gssearchParams += "-dBATCH"
        $gssearchParams += "-dEncryptionR#3"
        $gssearchParams += "-dKeyLength#128"
        $gssearchParams += "-sDEVICE=pdfwrite"
        $gssearchParams += "-dPDFSETTINGS=/default"
        $gssearchParams += "-dPermissions#$($pdfprotect)"
        $gssearchParams += "-sOwnerPassword=$($pdfpassword)"
        $gssearchParams += "-dCompatibilityLevel=1.4"
        $gssearchParams += "-sOutputFile=""$($pdffile)"""
        $gssearchParams += """$psfile"""
    
    #Start time measurement - Ghostscript
        $tempgstimestart = Get-Date
    
    #Invoke GhostScript with parameters
        $erg = Start-Process $gspath $gssearchParams -Wait -NoNewWindow -PassThru 2>>$pdf2smtp_ps_error_log
    
    #Log processing time - Ghostscript
        $tempgstimeend = get-date
        $tempgstime = "$(($tempgstimeend - $tempgstimestart).Minutes):$(($tempgstimeend - $tempgstimestart).Seconds)"
        "PDF2SMTP_PS $(Get-Date) Ghostscript processing time: $($tempgstime) Minute(s):Second(s)" >> $pdf2smtp_ps_log
        
        if($debug)
            { "PDF2SMTP_PS $(Get-Date) DEBUG $($gssearchParams)">>$pdf2smtp_ps_log }
        
        if(($erg.ExitCode) -ne 0)
        {
            "PDF2SMTP_PS $(Get-Date) ERROR: Ghostscript ExitCode: $($erg.ExitCode). Input-File: $($psfile)" >> $pdf2smtp_ps_log
            if($errormails)
            { Send-MailMessage -SmtpServer $MailSmtpServer -UseSsl -Subject "PDF2SMTP ERROR: Ghostscript" -Body "ERROR: Ghostscript ExitCode: $($erg.ExitCode)" -from $MailFrom -To $ErrorMailsTo -ErrorAction SilentlyContinue }
        }
        else
        {
            #Ghostscript processing successful
            #Komprimierungsvorgang pruefen 0Byte...???ß
                if($debug)
                    { "PDF2SMTP_PS $(Get-Date) DEBUG Ghostscript ExitCode: $($erg.ExitCode). Input-File: $($psfile)">>$pdf2smtp_ps_log }
            
                if(Test-Path $pdffile) 
                {
                    $pdffilecreatonsuccess = $true
                    


                    #Possible further pdffile manipulation, for example add watermark or text stamp. Also printername dependend possible
                        #if($env:REDMON_PRINTER -eq "PRINTERNAME") {}


                    
                    #LOG
                        "PDF2SMTP_PS $(Get-Date) Start prepare Mail">> $pdf2smtp_ps_log

                    #Prepare Mail
                        #MailTo
                            $tempMailTo = (($env:REDMON_USER).Trim())
                            #Find User Mail-Address according to its Active Directory samaccountname
                                $searcher = [adsisearcher]"(samaccountname=$($tempMailTo))"
                                $MailTo = $searcher.FindOne().Properties.mail
                
                        #Mail Subject
                            $MailSubject = $env:REDMON_DOCNAME
                
                        #Build Mail-Body Text
                            if($env:REDMON_PRINTER -eq "SPECIALPRINTERNAME")
                            { $MailBody = "Your SPECIAL conversion job has been executed`n" }           # Special printer success message
                            else
                            { $MailBody = "Your conversion job has been executed.`n" }           # Standard print to PDF success message

                            $MailBody += "---------------------------------------------------`n"
                            $MailBody += "Printout generated by:       $($env:REDMON_USER)`n"       # User wich printed
                            #$MailBody += "Computer:		    $($env:REDMON_MACHINE)`n"          # Computer on which the printjob was startet by the user (not the printserver computername)
                            $MailBody += "Printer: 		    $($env:REDMON_PRINTER)`n"              # Name of the printer used
                            $MailBody += "Document name: 	    $($env:REDMON_DOCNAME)`n"          # Name of the Document which has been printed
                            $MailBody += "Jobnumber:		    $($env:REDMON_JOB)`n"              # Number of the redmon jobnumber
                            $MailBody += "---------------------------------------------------`n"
                            $MailBody += "Created by $($env:REDMON_PRINTER) Version 1.2 at $(Get-Date)`n"
                
                        #Mail Attachment
                            $MailAttachment = $pdffile

                        if($debug)
                        { "PDF2SMTP_PS $(Get-Date) DEBUG .$($tempMailTo). .$($MailTo). .$($MailSubject). .$($MailBody). .$($MailAttachment). .$($env:REDMON_JOB)." >> $pdf2smtp_ps_log}


                    #LOG
                        "PDF2SMTP_PS $(Get-Date) Start send Mail">> $pdf2smtp_ps_log

                    #Send Mail
                        if($MailTo)
                        {
                            #Start time measurement - sendmail
                                $tempsendmailtimestart = Get-Date

                            try 
                            {
                                #Send mail Main-Server
                                    if($env:REDMON_PRINTER -eq "SPECIALPRINTERNAME")
                                    { $sendmailstatus = Send-MailMessage -SmtpServer $MailSmtpServer -UseSsl -Subject "$($MailSubject)" -Body "$($MailBody)" -from $MailFrom -To $MailTo -Encoding ([System.Text.Encoding]::UTF8) }
                                    else
                                    { $sendmailstatus = Send-MailMessage -SmtpServer $MailSmtpServer -UseSsl -Subject "$($MailSubject)" -Body "$($MailBody)" -from $MailFrom -To $MailTo -Attachments "$($MailAttachment)" -Encoding ([System.Text.Encoding]::UTF8) }
                                if($?)
                                {
                                    $mailsentsuccess = $true
                                    "PDF2SMTP_PS $(Get-Date) Mail sent to $($MailTo)" >> $pdf2smtp_ps_log
                                }
                                else
                                { 
                                    $mailsentsuccess = $false
                                    "PDF2SMTP_PS $(Get-Date) ERROR: sending Mail error1: $($?) - Possible SPAM Block?" >> $pdf2smtp_ps_log
                                    #Check SPAM Block List
                                    foreach ($server in $blocklistServers)
                                    {
                                        $fqdn = "$reversedSendingMailIP.$server"
                                        if(Resolve-DnsName $fqdn -ErrorAction SilentlyContinue)
                                        {
                                            "PDF2SMTP_PS $(Get-Date) ERROR: possible SPAM-Block for IP $($SendingMailSourceIP) on $($server)" >> $pdf2smtp_ps_log
                                        }
                                    }
                                    #Send mail Failover-Server
                                        $pw = $MailPwFailover | ConvertTo-SecureString -AsPlainText -Force
                                        $cred = New-Object System.Management.Automation.PSCredential $MailuserFailover, $pw
                                        $sendmailstatus = Send-MailMessage -SmtpServer $MailSmtpServerFailover -UseSsl -Subject "$($MailSubject)" -Body "$($MailBody)" -from $MailFrom -To $MailTo -Attachments "$($MailAttachment)" -Credential $cred -Encoding ([System.Text.Encoding]::UTF8)
                                    if($?)
                                    {
                                        $mailsentsuccess = $true
                                        "PDF2SMTP_PS $(Get-Date) Mail sent to $($MailTo) - Failover" >> $pdf2smtp_ps_log
                                    }
                                    else
                                    { 
                                        $mailsentsuccess = $false
                                        "PDF2SMTP_PS $(Get-Date) ERROR: sending Mail Failover error1: $($?)" >> $pdf2smtp_ps_log
                                    }
                                }
                            }
                            catch 
                            {
                                "PDF2SMTP_PS $(Get-Date) ERROR: sending Mail error2: $($?)" >> $pdf2smtp_ps_log
                                $mailsentsuccess = $false
                            }
                    
                            #Log processing time - sendmail
                                $tempsendmailtimeend = get-date
                                $tempsendmailtime = "$(($tempsendmailtimeend - $tempsendmailtimestart).Minutes):$(($tempsendmailtimeend - $tempsendmailtimestart).Seconds)"
                                "PDF2SMTP_PS $(Get-Date) Send Mail processing time: $($tempsendmailtime)  Minute(s):Second(s)" >> $pdf2smtp_ps_log
                        }
                        else
                        { 
                            "PDF2SMTP_PS $(Get-Date) ERROR: MailTo empty: User $($tempMailTo) - $($env:REDMON_USER) not found!" >> $pdf2smtp_ps_log
                            if($errormails)
                                { Send-MailMessage -SmtpServer $MailSmtpServer -UseSsl -Subject "PDF2SMTP_PS ERROR: Mail" -Body "ERROR: MailTo empty: User $($tempMailTo) - $($env:REDMON_USER) not found!" -from $MailFrom -To $ErrorMailsTo -ErrorAction SilentlyContinue }
                        }
        }
                else
                {
                    "PDF2SMTP_PS $(Get-Date) ERROR: PDFFile $($pdffile) not found!">> $pdf2smtp_ps_log
                    if($errormails)
                        { Send-MailMessage -SmtpServer $MailSmtpServer -UseSsl -Subject "PDF2SMTP_PS $(Get-Date) ERROR: PDFFile" -Body "ERROR: PDFFile $($pdffile) not found!" -from $MailFrom -To $ErrorMailsTo -ErrorAction SilentlyContinue }

                    $pdffilecreatonsuccess = $false
                }
        }
}
else
{
    "PDF2SMTP_PS $(Get-Date) ERROR: PSFile $($psfile) not found!">> $pdf2smtp_ps_log
    if($errormails)
        { Send-MailMessage -SmtpServer $MailSmtpServer -UseSsl -Subject "PDF2SMTP_PS $(Get-Date) ERROR: PSFile" -Body "ERROR: PSFile $($psfile) not found!" -from $MailFrom -To $ErrorMailsTo -ErrorAction SilentlyContinue }
    $psfilecreatonsuccess = $false
}

#LOG
    "PDF2SMTP_PS $(Get-Date) Start deleting temp files">> $pdf2smtp_ps_log

#Deleting temp files
    try
    { Remove-Item -Path $pdffile}
    catch 
    {
        #LOG
            "PDF2SMTP_PS $(Get-Date) Error: error deleating temp PDF: $($pdffile)">> $pdf2smtp_ps_log
    }
    try
    { Remove-Item -Path $psfile }
    catch 
    {
        #LOG
            "PDF2SMTP_PS $(Get-Date) ERROR: error deleating temp PDF: $($psfile)">> $pdf2smtp_ps_log
    }


#Log processing time
    $temptimeend = get-date
    $temptime = "$(($temptimeend - $temptimestart).Minutes):$(($temptimeend - $temptimestart).Seconds)"
    "PDF2SMTP_PS $(Get-Date) Entire processing time: $($temptime) Minute(s):Second(s)" >> $pdf2smtp_ps_log


#DEBUG Send Time-Mails
if($debug)
{
    $sendmailtime = $false
    if(($tempsendmailtimeend - $tempsendmailtimestart).Minutes -gt 1)
        { $sendmailtime = $true }
    if(($tempgstimeend - $tempgstimestart).Minutes -gt 1)
        { $sendmailtime = $true }
    if(($tempredfiletimeend - $tempredfiletimestart).Minutes -gt 1)
        { $sendmailtime = $true }
    if(($temptimeend - $temptimestart).Minutes -gt 1)
        { $sendmailtime = $true }
    if($sendmailtime)
    {
        $sendtimeMailSubject = "Time processing Redmon Job $($env:REDMON_JOB)"

        $sendtimeMailBody =  "Time processing for redmon job$($env:REDMON_JOB) - long duration.`n"
        $sendtimeMailBody += "---------------------------------------------------`n"
        $sendtimeMailBody += "Redfile Time:                      $($tempredfiletime)  Minute(s):Second(s)`n"
        $sendtimeMailBody += "Ghostscript Time:              $($tempgstime)  Minute(s):Second(s)`n"
        $sendtimeMailBody += "SendMail Time:                  $($tempsendmailtime)  Minute(s):Second(s)`n"
        $sendtimeMailBody += "Processing Time:                $($temptime)  Minute(s):Second(s)`n"
        $sendtimeMailBody += "REDMON_DOCNAME:	    $($env:REDMON_DOCNAME)`n"
        $sendtimeMailBody += "REDMON_USER:	    $($env:REDMON_USER)`n"
        $sendtimeMailBody += "REDMON_Printer:	    $($env:REDMON_PRINTER)`n"
        $sendtimeMailBody += "PS file creation:                   $($psfilecreatonsuccess)`n"
        $sendtimeMailBody += "PDF file creation:                $($pdffilecreatonsuccess)`n"
        $sendtimeMailBody += "Mail sent:                             $($mailsentsuccess)`n"
        $sendtimeMailBody += "---------------------------------------------------`n"
        $sendtimeMailBody += "Created by PDF2SMTP_PS Version 1.0 at $(Get-Date)`n"

                    try 
                    {
                        #$sendmailstatus = 
                        Send-MailMessage -SmtpServer $MailSmtpServer -UseSsl -Subject "$($sendtimeMailSubject)" -Body "$($sendtimeMailBody)" -from $MailFrom -To $sendtimemailsmailto -Encoding ([System.Text.Encoding]::UTF8) -ErrorAction SilentlyContinue
                        if($?)
                        {
                            $mailsentsuccess = $true
                            "PDF2SMTP_PS $(Get-Date) Mail long processing time sent to $($MailTo)" >> $pdf2smtp_ps_log
                        }
                        else
                        { 
                            $mailsentsuccess = $false
                            "PDF2SMTP_PS $(Get-Date) ERROR: sending Mail long processing time error1: $($?)" >> $pdf2smtp_ps_log
                        }
                    }
                    catch 
                    {
                        "PDF2SMTP_PS $(Get-Date) ERROR: sending Mail long processing time error2: $($?)" >> $pdf2smtp_ps_log
                        $mailsentsuccess = $false
                    }
        "PDF2SMTP_PS $(Get-Date) DEBUG slow processing time: -SmtpServer $($MailSmtpServer) -UseSsl -Subject sendtimeMailSubject -Body sendtimeMailBody -from $($MailFrom) -To $($sendtimemailsmailto)" >> $pdf2smtp_ps_log
        "PDF2SMTP_PS $(Get-Date) DEBUG slow processing time: $($temptime) Second(s)" >> $pdf2smtp_ps_log
    }
}

#LOG
    "PDF2SMTP_PS $(Get-Date) -----------------------------------">> $pdf2smtp_ps_log
    "PDF2SMTP_PS $(Get-Date) END PDF2SMTP_PS">> $pdf2smtp_ps_log
    "PDF2SMTP_PS $(Get-Date) -----------------------------------">> $pdf2smtp_ps_log
