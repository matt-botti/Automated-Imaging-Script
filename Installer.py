import sys
import time
from pywinauto import Application

confirmPrinter = False
confirmOffice = False
confirmOfficePK = False
confirmSophos = False
confirmZoom = False

#parse cmd args here to change flags
if "printer" in sys.argv: confirmPrinter = True
if "office" in sys.argv: confirmOffice = True
if "officePK" in sys.argv: confirmOfficePK = True
if "sophos" in sys.argv: confirmSophos = True
if "zoom" in sys.argv: confirmZoom = True

def printerInstall():
    # Need to launch printer window in PowerShell before running this script
    print("Beginning printer installation")
    app = Application(backend='uia').connect(title="Add a device")
    printer = app.window()

    printer.FollowYouPrinterOnEquitracButton.wait('ready exists', timeout=60).click()
    printer.NextButton.wait('ready exists', timeout=30).click()

    printer = app.window(title="Add Printer")
    printer.FinishButton.wait('ready exists',timeout=60).click()
    
    
def officeInstall():
    print("Launching Office 2010 installer")
    app = Application(backend='uia').start("//***REMOVED***/setup apps/office 2010/setup.exe", timeout=300)   
    office = app.window(title="Microsoft Office Professional Plus 2010")

    #temp solution to check the box to agree to TOS
    if(office.GroupBox2.CheckBox.exists(timeout=300)):
        print("Agreeing to TOS...")
        office.GroupBox2.CheckBox.wait('ready exists',timeout=300)
        #Sends Alt+a keystroke to program
        office.type_keys("%a")
        office.ContinueButton.wait('ready exists',timeout=300).click()
    #click the 'install now' button
    print("Clicking Install Now button...")
    office.InstallNow.wait('ready exists', timeout=300).click()
    #wait for progress bar to show up, then wait for it to finish
    print("Waiting for progress bar...")
    office.Progress.wait('exists',timeout=300)
    office.Progress.wait_not('exists',timeout=600)
    #then finally exit the installer
    print("Exiting the installer...")
    office.CloseButton.wait('exists ready',timeout=600).click()
    #TODO - vv This doesn't work. Need to figure out how to hit "no" on the popup window after closing the main window
    office = app.window(title="Setup")
    if(office.NoButton.exists()):
        office.NoButton.wait('exists ready', timeout=300).click()
    print("")
        
        
def officePK():
    print("Launching Office 2010 installer for product key...")
    app = Application(backend="uia").start("//***REMOVED***/setup apps/office 2010/setup.exe",timeout=30)
    #app = Application(backend="uia").connect(title="Microsoft Office Professional Plus 2010")
    #start office setup
    office = app.window(title="Microsoft Office Professional Plus 2010")

    #First menu - Click 'Enter Product Key' radio button, then continue button
    print("Selecting the Product Key radio button and hitting Continue...")
    office.ProductKey.click()
    office.ContinueButton.click()
    #Enter the product key, wait for the buttons to become clickable, then click them
    print("Entering product key and hitting continue...")
    office.wait('ready').type_keys("RKFCH-HKWJM-WM4RK-BB6YW-4BWYF")
    office.GroupBoxEnterYourProductKey.ContinueButton.wait('ready',timeout=30).click()
    office.ChooseTheInstallationYouWantGroupBox.ContinueButton.wait('ready').click()
    #first wait for the progress bar to exist
    print("Waiting for progress bar...")
    office.Progress.wait('exists',timeout=30)
    #then wait for it to finish (not exist)
    office.Progress.wait_not('exists',timeout=600)
    office.CloseButton.wait('exists ready',timeout=600).click()
    print("Exiting the installer...")
    #TODO - vv This doesn't work. Need to figure out how to hit "no" on the popup window after closing the main window
    office = app.window(title="Setup")
    if(office.NoButton.exists()):
        office.NoButton.wait('exists ready', timeout=300).click()
    print("")

def sophosInstall():
    print("Launching Sophos installer...")
    app = Application(backend='uia').start("//***REMOVED***/setup apps/Sophos Central Endpoint Installer/SophosSetup.exe")
    time.sleep(5)
    app = Application(backend="uia").connect(title="Sophos Install")
    sophos = app.window()

    print("Clicking the Continue button...")
    sophos.ContinuePane.wait('ready exists',timeout=600).click_input()
    print("Clicking the Install button...")
    sophos.InstallPane.wait('ready exists',timeout=600).click_input()
    print("Waiting for progress bar...")
    sophos.InstallationWillTakeAbout10Minutes.wait('exists',timeout=600)
    sophos.InstallationWillTakeAbout10Minutes.wait_not('exists',timeout=1200)
    #Restart checkbox sometimes does not appear.
    if(sophos.Finish.exists(timeout=60)):
        if(sophos.RestartMyComputerNow.exists()):
            print("Unchecking restart checkbox, then exiting the installer...")
            sophos.RestartMyComputerNow.wait('ready exists').click_input(coords=(0,0))
    print("Closing Sophos installer")
    sophos.Finish.wait('ready exists',timeout=60).click_input()
    print("")
    
    
def zoomInstall():
    print("Launching Zoom installer...")
    app = Application(backend='uia').start("//***REMOVED***/setup apps/zoom/zoominstaller.exe",timeout=300)
    app = Application(backend='uia').connect(title="Zoom",timeout=600)
    zoom = app.window(title="Zoom")

    print("Pressing the close button...")
    zoom.Done.wait('ready exists',timeout=300).click()
    zoom = app.window(title="Zoom Cloud Meetings")
    print("Pressing the other close button...")
    zoom.CloseButton.wait('ready exists',timeout=60).click()
    print("")


if (confirmPrinter):
    printerInstall()
if (confirmOffice):
    officeInstall()
if (confirmOfficePK):
    officePK()
if (confirmSophos):
    sophosInstall()
if (confirmZoom):
    zoomInstall()
