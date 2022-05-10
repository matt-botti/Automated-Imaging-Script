$confirmAll = Read-Host "Do you want to install Office 2010, Sophos, and Zoom? If no, you will be prompted to individually select which programs to install. (y/n)"
while (($confirmAll -notmatch "[yYnN]") -or ($confirmAll -notmatch "yyyy")) {
    $confirmAll = Read-Host "Please answer with either y or n. Do you want to install Office 2010, Sophos, Impero, and Zoom? (y/n)"
}
if (($confirmAll -match "[yY]") -or ($confirmAll -match "yyyy")) {
    $confirmOffice = "y"
    $confirmSophos = "y"
    $confirmImpero = "y"
    $confirmZoom = "y"
} else {
    $confirmOffice = Read-Host "Do you want Office 2010 installed? (y/n)"
    while ($confirmOffice -notmatch "[yYnN]") {
        $confirmOffice = Read-Host "Please answer with either y or n. Do you want Office 2010 installed? (y/n)"
    }

    $confirmSophos = Read-Host "Do you want Sophos installed? (y/n)"
    while ($confirmSophos -notmatch "[yYnN]") {
        $confirmSophos = Read-Host "Please answer with either y or n. Do you want Sophos installed? (y/n)"
    }

    $confirmImpero = Read-Host "Do you want Impero installed? (y/n)"
    while ($confirmImpero -notmatch "[yYnN]") {
        $confirmImpero = Read-Host "Please answer with either y or n. Do you want Impero installed? (y/n)"
    }

    $confirmZoom = Read-Host "Do you want Zoom installed? (y/n)"
    while ($confirmZoom -notmatch "[yYnN]") {
        $confirmZoom = Read-Host "Please answer with either y or n. Do you want Zoom installed? (y/n)"
    }
}

Write-Output "Checking if Equitrac printer needs added..."
$printer = Get-Printer
if ($printer -like "*Follow You*") {
    Write-Output("Equitrac printer already installed on this machine")
    $confirmPrinter = "n"
} else {
    Write-Output "The Equitrac printer will be installed"
    $confirmPrinter = "y"
}

if ($confirmAll -notmatch "yyyy") {
    Write-Output " "

    Write-Output "Disabling Windows factory reset"
    reagentc.exe /disable

    Write-Output " "

    Write-Output "Setting timezone to Eastern Standard Time"
    Set-TimeZone "Eastern Standard Time"
    Write-Output "Synching time with server"
    W32tm /resync /force
    Write-Output " "

    Write-Output "Disabling Windows Firewall"
    Set-NetFirewallProfile -Profile Domain, Private, Public -Enabled False
    Write-Output " "

    Write-Output "Checking if SMB 1.0 is enabled..."
    $smb = Get-WindowsOptionalFeature -Online -FeatureName SMB1Protocol | Select-Object State
    if ($smb -like "*Enabled*") {
        Write-Output "SMB 1.0 is already enabled"
    } else {
        Write-Output "Enabling SMB 1.0"
        Enable-WindowsOptionalFeature -Online -NoRestart -FeatureName smb1protocol
    }



    Write-Output "Removing Xbox Game Bar"
    Get-AppxPackage Microsoft.XboxGamingOverlay | Remove-AppxPackage
    Write-Output " "

    Write-Output "Connecting to Wi-Fi"
    $SSID = netsh wlan show interfaces | select-string SSID
    if ($SSID -like "*ACADEMYSCHOOLS*") {
        Write-Output "Already connected to ACADEMYSCHOOLS"
    } else {
        Netsh WLAN delete profile "ACADEMYSCHOOLS"
        Netsh WLAN add profile filename="\\***REMOVED***\setup apps\1 Fresh Image Script\WLAN-ACADEMYSCHOOLS.XML"
        Netsh WLAN connect name="ACADEMYSCHOOLS"
        Write-Output "Done"
    }
    Write-Output " "

    Write-Output "Checking Windows activation"
    $LicenseStatus = Get-CimInstance SoftwareLicensingProduct -Filter "Name like 'Windows%'" | Where-Object { $_.PartialProductKey } | Select-Object LicenseStatus

    if ($LicenseStatus -like '*1*') {
        Write-Output "Windows is already activated"
    } else {
        slmgr.vbs /ipk WNYMG-QRJK7-268GK-X3CJ7-BDWXM
        Write-Output "Windows is now activated"
    }
    Write-Output " "
}

if ($confirmOffice -match "[yY]") {
    Write-Output "Checking if Office 2010 is installed"
    $officeInstalled = $null -ne (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object { $_.DisplayName -eq "Microsoft Office Professional Plus 2010" })
    if ($officeInstalled) {
        Write-Output "Office 2010 is already installed"
        $confirmOffice = "n"
    } else {
        Write-Output "Office 2010 will be installed"
    }   
    Write-Output "Checking Office activation status..."
    $OfficeActivation = cscript "C:\Program Files\Microsoft Office\Office14\OSPP.vbs" /dstatus | select-string "LICENSE STATUS"
    if ($OfficeActivation -like "*---LICENSED---*") {
        Write-Output "Office is already activated"
        $confirmOfficePK = "n"
    } else {
        Write-Output "Office product key will be installed"
        $confirmOfficePK = "y"    
    }

} else {
    Write-Output "You have elected not to install Office 2010"
}

Write-Output " "


if ($confirmSophos -match "[yY]") {
    $sophosInstalled = $null -ne (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object { $_.DisplayName -eq "Sophos Endpoint Agent" })
    if ($sophosInstalled) {
        Write-Output "Sophos is already installed"
        $confirmSophos = "n"
    } else {
        Write-Output "Sophos will be installed"
    }

} else {
    Write-Output "You have elected not to install Sophos"
}

if ($confirmImpero -match "[yY]") {
    $imperoInstalled = $null -ne (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object { $_.DisplayName -eq "Impero Client" })
    if ($imperoInstalled) {
        Write-Output "Impero is already installed"
        $confirmImpero = "n"
    } else {
        Write-Output "Running Impero installer..."
        msiexec.exe /i "\\***REMOVED***\Setup Apps\Impero\Impero Education Pro v8503 Installers\ImperoClientSetup8503.msi" /passive /norestart
        Write-Output "Creating txt file to point Impero at its server..."
        New-Item "C:\Program Files (x86)\Impero Solutions Ltd\Impero Client\ServerIPFixed.txt"
        Set-Content "C:\Program Files (x86)\Impero Solutions Ltd\Impero Client\ServerIPFixed.txt" "***REMOVED***"
    }

} else {
    Write-Output "You have elected not to install Impero"
}

Write-Output " "
if ($confirmZoom -match "[yY]") {
    $zoomInstalled = $null -ne (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object { $_.DisplayName -eq "Zoom" })
    if ($zoomInstalled) {
        Write-Output "Zoom is already installed"
        $confirmZoom = "n"
    } else {
        Write-Output "Zoom will be installed"
    }

    Write-Output " "
} else {
    Write-Output "You have elected not to install Zoom"
}

#construct command line argument for python script
$pythonArgs = ""
if ($confirmPrinter -match "[yY]") {
    $pythonArgs += " printer"
    rundll32 printui.dll, PrintUIEntry /il 
}
if ($confirmOffice -match "[yY]") {
    $pythonArgs += " office"
    if ($confirmOfficePK -match "[yY]") {
        $pythonArgs += " officePK"
    }
}
if ($confirmSophos -match "[yY]") {
    $pythonArgs += " sophos"
}
if ($confirmZoom -match "[yY]") {
    $pythonArgs += " zoom"
}

#if we aren't going to do anything with the python script, don't launch it
if (($confirmPrinter -match "[nN]") -and ($confirmOffice -match "[nN]") -and ($confirmOfficePK -match "[nN]") -and ($confirmSophos -match "[nN]") -and ($confirmZoom -match "[nN]")) {
    Write-Output "Python script is not needed this time"
} else {
    Start-Process "\\***REMOVED***\Setup Apps\2 Fresh Image Automated Script\dist\Installer.exe" $pythonArgs -Wait -NoNewWindow
}

Write-Output "Opening Windows Update"
control update
Write-Output " "

Write-Output "Don't forget to change the OU and restart!"
hostname

Read-Host -Prompt "Press Enter to exit"
