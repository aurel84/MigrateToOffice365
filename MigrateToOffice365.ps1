<#
Script to perform automated migration, or installation of Office365.
Performs an initial check to see if user has cab files already present in C:\Temp\Office
If it does not, downloads them from Microsoft CDN.  After this step, script checks to 
see what current Office 2016 applciations User has, then performs the recommended
install using the appropriate .xml file as a reference.

By John Hawkins | johnhawkins3d@gmail.com
#>

### Variables ###
$logFile = ( ($env:computername | Select-Object) + "_" + "InstallOffice365ProPlus" + "_" + (Get-Date -Format "MM_dd_yyyy") ) # Name of logfile that is stored in C:\Temp

### Object ###
$officeApplications=[PSCustomObject]@{

    Office365 = (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\O365*") # Test for office365 proplus
    Visio16 = (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Office16.VIS*") # Test for visio 2016
    Project16 = (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Office16.PRJ*") # Test for project 2016

}

### Functions ##
function  RestartExplorer {

    Write-Host "Restarting Explorer."

    taskkill /f /im explorer.exe

    Start-Process explorer.exe

}

function PromptUser {

    Write-Host "Giving user notice of installation instructions."

    $wshell0 = New-Object -ComObject Wscript.Shell
    $wshell0.Popup("Microsoft Office365 ProPlus has finished downloading.  
The installation of Office365 will require you to exit all Office Applications.  When installation is completed, you will need to re-pin your shortcuts for Office Applications to the Start Menu and Taskbar.",0,"IT Notice",0x0)

}

### Start of Log file for script ###
Start-Transcript -path "C:\Temp\$logFile.txt"

Write-Host "Checking for Data files for Office365 in C:\Temp\Office..."

if ( -not ( Test-Path "C:\Temp\Office\*" ) ) {

    Write-Host "No Data files for Office365 found, downloading them from CDN."

    C:\Temp\OfficeLocal\setup.exe /download C:\Temp\OfficeLocal\Office365ProjectVisioProPlus.xml

}   else    {

    Write-Output "Data files for Office365 are present, download from CDN not needed."

}

if ( $officeApplications.Office365 -eq $true )   {

    Write-Host "Office365ProPlus already installed, no further action needed."

}   else    {

    ### Foreach statement to run through the different conditions on which to install or migrate user to Office365 Proplus and Applications ###
    Write-Host "Checking for conditions what Applications are needed to be installed..."

    ForEach ( $App in $officeApplications ) {

        if  ( $App.Visio16 -eq $false -and $App.Project16  -eq $false -and $App.Office365 -eq $false )  {

        Write-Host "Detetected Office 2016.  Migrating user to Office365 Pro Plus."

        PromptUser

        C:\Temp\OfficeLocal\setup.exe /configure C:\Temp\OfficeLocal\Office365ProPlus.xml

        RestartExplorer

        }   elseif ( $App.Visio16 -eq $true -and $App.Project16  -eq $false -and $App.Office365 -eq $false )   {

        Write-Host "Detected Office 2016 and Visio 2016.  Migrating user to Office365 ProPlus and Visio."

        PromptUser

        C:\Temp\OfficeLocal\setup.exe /configure C:\Temp\OfficeLocal\Office365VisioProPlus.xml

        RestartExplorer

        }   elseif ( $App.Visio16 -eq $false -and $App.Project16  -eq $true -and $App.Office365 -eq $false )   {

        Write-Host "Detected Office 2016 and Project 2016.  Migrating user to Office365 ProPlus and Project."

        PromptUser

        C:\Temp\OfficeLocal\setup.exe /configure C:\Temp\OfficeLocal\Office365ProjectProPlus.xml

        RestartExplorer

        }   elseif ( $App.Visio16 -eq $true -and $App.Project16  -eq $true -and $App.Office365 -eq $false )   {

        Write-Host "Detected Office 2016, Visio 2016, and Project 2016.  Migrating user to Office365 ProPlus, Visio, and Project."

        PromptUser

        C:\Temp\OfficeLocal\setup.exe /configure C:\Temp\OfficeLocal\Office365ProjectVisioProPlus.xml

        RestartExplorer
        
        }

    }

}

### End of Log file for script ###
Stop-Transcript

Exit
