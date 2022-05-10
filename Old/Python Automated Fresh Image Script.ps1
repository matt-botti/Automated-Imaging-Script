$confirmAll = Read-Host -Prompt "Do you want to install Office 2010, Sophos, and Zoom? If no, you will be prompted to individually select which programs to install. (y/n)"
while ($confirmAll -notmatch "[yYnN]"){
    $confirmAll = Read-Host -Prompt "Please answer with either y or n. Do you want to install Office 2010, Sophos, and Zoom? (y/n)"
}
if ($confirmAll -match "[yY]") {
    $confirmOffice = "y"
    $confirmSophos = "y"
    $confirmZoom = "y"
} else {
    $confirmOffice = Read-Host -Prompt "Do you want Office 2010 installed? (y/n)"
    while ($confirmOffice -notmatch "[yYnN]"){
       $confirmOffice = Read-Host -Prompt "Please answer with either y or n. Do you want Office 2010 installed? (y/n)"
    }

    $confirmSophos = Read-Host -Prompt "Do you want Sophos installed? (y/n)"
    while ($confirmSophos -notmatch "[yYnN]"){
        $confirmSophos = Read-Host -Prompt "Please answer with either y or n. Do you want Sophos installed? (y/n)"
    }

    $confirmZoom = Read-Host -Prompt "Do you want Zoom installed? (y/n)"
    while ($confirmZoom -notmatch "[yYnN]"){
        $confirmZoom = Read-Host -Prompt "Please answer with either y or n. Do you want Zoom installed? (y/n)"
    }
}
#Make sure python is installed
Write-Output "This is the version of the script that requires Python is installed. Checking if Python is installed..."
$pythonInstalled = $null -ne (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object { $_.DisplayName -eq "Python 3.10.4 Core Interpreter (64-bit)" })
if($pythonInstalled){
    Write-Output "Python is already installed"
} else {
    Write-Output "Running Python 3.10.4 installer..."
    Start-Process "\\***REMOVED***\setup apps\Python\python-3.10.4 64-bit.exe" -Wait
}

Write-Output " "

Write-Output "Disabling Windows factory reset"
reagentc.exe /disable

Write-Output "Add Equitrac printer using this prompt to ensure drivers get installed"
rundll32 printui.dll,PrintUIEntry /il 
Write-Output " "

Write-Output "Setting timezone to Eastern Standard Time"
Set-TimeZone "Eastern Standard Time"
Write-Output "Synching time with server"
W32tm /resync /force
Write-Output " "

Write-Output "Disabling Windows Firewall"
Set-NetFirewallProfile -Profile Domain,Private,Public -Enabled False
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

if ($confirmOffice -match "[yY]"){
    Write-Output "Checking if Office 2010 is installed"
    $officeInstalled = $null -ne (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object { $_.DisplayName -eq "Microsoft Office Professional Plus 2010" })
    if ($officeInstalled){
        Write-Output "Office 2010 is already installed"
       
    } else {
        Write-Output "Installing Office 2010..."
        Py "\\***REMOVED***\setup apps\2 Fresh Image Automated Script\OfficeInstallers\OfficeInstall\OfficeInstall.py" -Wait
    }   
    Write-Output "Checking Office activation status..."
    $OfficeActivation = cscript "C:\Program Files\Microsoft Office\Office14\OSPP.vbs" /dstatus | select-string "LICENSE STATUS"
    if ($OfficeActivation -like "*---LICENSED---*") {
        Write-Output "Office is already activated"
    } else {
        Write-Output "Installing Office product key..." 
        Py "\\***REMOVED***\setup apps\2 Fresh Image Automated Script\OfficeInstallers\OfficeProductKey\OfficeProductKey.py" -Wait     
    }

} else {
    Write-Output "You have elected not to install Office 2010"
}

Write-Output " "


if ($confirmSophos -match "[yY]"){
    $sophosInstalled = $null -ne (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object { $_.DisplayName -eq "Sophos Endpoint Agent" })
    if($sophosInstalled){
        Write-Output "Sophos is already installed"
        
    } else {
        Write-Output "Running Sophos Installer..."
        Py "\\***REMOVED***\setup apps\2 Fresh Image Automated Script\SophosInstaller\SophosInstall.py" -Wait
    }

} else {
    Write-Output "You have elected not to install Sophos"
}

Write-Output " "
if ($confirmZoom -match "[yY]"){
    $zoomInstalled = $null -ne (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object { $_.DisplayName -eq "Zoom" })
    if($zoomInstalled){
        Write-Output "Zoom is already installed"

    } else {
        Write-Output "Running Zoom Installer"
        Py "\\***REMOVED***\setup apps\2 Fresh Image Automated Script\ZoomInstaller\ZoomInstall.py" -Wait
    }

    Write-Output " "
} else {
    Write-Output "You have elected not to install Zoom"
}

Write-Output "Opening Windows Update"
control update
Write-Output " "

Write-Output "Don't forget to change the OU, push Impero, and restart!"
hostname
Write-Output "And uninstall Python!"
Read-Host -Prompt "Press Enter to exit"
