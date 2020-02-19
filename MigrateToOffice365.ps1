<#
Script to perform automated migration and / or installation of Office365.
Performs an initial check to see if user has cab files already present in C:\Temp\Office
If it does not, downloads them from Microsoft CDN.  After this step, script checks to 
see what current Office 2016 applciations User has, then performs the recommended
install using the appropriate .xml file as a reference.

By John A. Hawkins | johnhawkins3d@gmail.com
#>

### Variables ###
$logFile = ( ($env:computername | Select-Object) + "_" + "InstallOffice365ProPlus" + "_" + (Get-Date -Format "MM_dd_yyyy") ) # Name of logfile that is stored in C:\Temp


$officeApplications=[PSCustomObject]@{  # Custom object o create a table to tabulate what office programs user has. 

    Office16 = (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Office16.PROPLUS*") # Test for office 2016
    Visio16 = (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Office16.VSPRO") # Test for visio 2016
    Project16 = (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Office16.PRJPRO") # Test for project 2016
    Office365 = (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\O365*") # Test for office365 proplus

}

### Functions ##
function  restartExplorer {
    
    taskkill /f /im explorer.exe

    Start-Process explorer.exe

}

function promptUser {

    $wshell0 = New-Object -ComObject Wscript.Shell
    $wshell0.Popup("Microsoft Office365 ProPlus has finished downloading, and installation will require you to close and uninstall your version of Office.",0,"IT Notice",0x0)

}

# Start of Log file for script
Start-Transcript -path "C:\Temp\$logFile.txt"

# Downloads the cab files for Office365 and stores them in C:\Temp\Office\*
if ( -not ( Get-ItemProperty -Path $pathToOfficeReg | Select-Object "uninstallstring" | Select-String "OfficeClickToRun.exe" | Select-String "Office" | Select-String "O365ProPlusRetail" ) )    {

    if ( -not ( Test-Path "C:\Temp\Office\*" ) ) {

        C:\Temp\OfficeLocal\setup.exe /download C:\Temp\OfficeLocal\Office365ProjectVisioProPlus.xml

    }

    else    {

        Write-Output "Office Files are already installed."

    }

}

ForEach ($App in $officeApplications)   {

    if ( $App.Office16 -eq $true -and $App.Visio16 -eq $false -and $App.Project16  -eq $false )   {

        Write-Host "Install Office365 ProPlus."

        PromptUser

        C:\Temp\OfficeLocal\setup.exe /configure C:\Temp\OfficeLocal\Office365ProPlus.xml

        restartExplorer

    }

    elseif ( $App.Office16 -eq $true -and $App.Visio16 -eq $true -and $App.Project16  -eq $false )   {

        Write-Host "Install Office365 ProPlus and Visio."

        PromptUser

        C:\Temp\OfficeLocal\setup.exe /configure C:\Temp\OfficeLocal\Office365VisioProPlus.xml

        restartExplorer

    }

    elseif ( $App.Office16 -eq $true -and $App.Visio16 -eq $false -and $App.Project16  -eq $true )   {

        Write-Host "Install Office365 ProPlus and Project."

        PromptUser

        C:\Temp\OfficeLocal\setup.exe /configure C:\Temp\OfficeLocal\Office365ProjectProPlus.xml

        restartExplorer

    }
    
    elseif ( $App.Office16 -eq $true -and $App.Visio16 -eq $true -and $App.Project16  -eq $true )   {

        Write-Host "Install All Office 365 ProPlus, Visio, and Project."

        PromptUser

        C:\Temp\OfficeLocal\setup.exe /configure C:\Temp\OfficeLocal\Office365ProjectVisioProPlus.xml

        restartExplorer

    }

    elseif ( $App.Office365 -eq $true )   {

        Write-Host "Office365ProPlus already installed.  No further action needed."


    }

    else    {

        Write-Host "There was an error, check with this user."

    }

}

Stop-Transcript

Exit
