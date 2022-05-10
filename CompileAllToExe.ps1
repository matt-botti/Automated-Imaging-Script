ps2exe "\\***REMOVED***\setup apps\2 Fresh Image Automated Script\Automated Fresh Image Script.ps1" "\\***REMOVED***\setup apps\2 Fresh Image Automated Script\Automated Fresh Image Setup.exe"

<#After writing a Python script or program you can use PyInstaller
pip install pyinstaller 
And then do the following to create an exe file that should run without the need of having Python installed on the target PC
    pyinstaller --onefile program.py 
the ‘—onefile’ parameter is used to put all the needed files and libraries into a single executable file. 
#>

#pyinstaller.exe --onefile "\\***REMOVED***\Setup Apps\2 Fresh Image Automated Script\PrinterInstaller\PrinterInstaller.spec"
#pyinstaller.exe --onefile "\\***REMOVED***\setup apps\2 Fresh Image Automated Script\OfficeInstallers\OfficeInstall\OfficeInstall.spec"
#pyinstaller.exe --onefile "\\***REMOVED***\setup apps\2 Fresh Image Automated Script\OfficeInstallers\OfficeProductKey\OfficeProductKey.spec"
#pyinstaller.exe --onefile "\\***REMOVED***\setup apps\2 Fresh Image Automated Script\SophosInstaller\SophosInstall.spec"
#pyinstaller.exe --onefile "\\***REMOVED***\setup apps\2 Fresh Image Automated Script\ZoomInstaller\ZoomInstall.spec"

#pyinstaller.exe --onefile "\\***REMOVED***\setup apps\2 Fresh Image Automated Script\Installer.py"
# ^^ This is slower, but use if something goes wrong with this vv
pyinstaller.exe --onefile "\\***REMOVED***\setup apps\2 Fresh Image Automated Script\Installer.spec"