# pdf2smtp_ps
Print to networkprinter and get PDF-File via email on a Windows print server a PDF printer can be implemented using redmon / powershell / ghostscript which accepts documents and sends them to the user as PDF via e-mail (mailaddress is read from the Active Directory using the SAMAccountName)

Solution is based on the work of Frank Carius: https://www.msxfaq.de/tools/pdf2smtp/uebersicht.htm

Informations:

  Overview and detaild instructions (German):
  https://www.msxfaq.de/tools/pdf2smtp/uebersicht.htm
  Redmon:
  http://www.ghostgum.com.au/software/redmon.htm
  
 Instructions:
  Install ghostscript to c:\Program Files\gs\version\   for example "C:\Program Files\gs\gs9.52\bin\gswin64c.exe"
    ![Install-Folders](https://user-images.githubusercontent.com/21160938/120331654-4cb70580-c2ee-11eb-88f0-76be83c0b3ea.png)
  Install redmon to c:\Program Files\gs\redmon\
   see http://www.ghostgum.com.au/software/redmon19.htm#27

  Add Redmon Port
    RPT1: Properties
    
      Redirect this port to the program: C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe
      Arguments for this program are: -command C:\Skripte\pdf2smtp_ps.ps1
      Output: Program handles output
      Run: Normal
      Print Erros: Unchecked
      Run as User: Unchecked
      Shut down delay: 300 seconds
      
      Log Files
      Use Log File: Cecked
      Write standard output to the file: c:\Program Files\gs\redmon\pdf2smtp_redmon.log
      Debug: Unchecked
  Add Printer attached to Redmon Port
    Driver: HP Universal Printing PS
    Print to spooler, print after last page was spooled
    
    Share printer
  Make sure Active Directory Users mail attribut is set (Find User Mail-Address according to its Active Directory samaccountname)

  Save pdf2smtp_ps.ps1 to C:\Skripte\pdf2smtp_ps.ps1
