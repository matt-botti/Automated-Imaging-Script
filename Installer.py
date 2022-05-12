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

    printer.FollowYouPrinterOnEquitracButton.wait('ready exists', timeout=6000).click()
    printer.NextButton.wait('ready exists', timeout=6000).click()

    printer = app.window(title="Add Printer")
    printer.FinishButton.wait('ready exists',timeout=6000).click()
    
    
def officeInstall():
    print("Launching Office 2010 installer")
    app = Application(backend='uia').start("//***REMOVED***/setup apps/office 2010/setup.exe", timeout=6000)   
    office = app.window(title="Microsoft Office Professional Plus 2010")

    #check the box to agree to TOS
    if(office.GroupBox2.CheckBox.exists(timeout=6000)):
        print("Agreeing to TOS...")
        office.GroupBox2.CheckBox.wait('ready exists',timeout=6000)
        #Sends Alt+a keystroke to program, to proceed to the next screen
        office.type_keys("%a")
        office.ContinueButton.wait('ready exists',timeout=6000).click()
    #click the 'install now' button
    print("Clicking Install Now button...")
    office.InstallNow.wait('ready exists', timeout=6000).click()
    #wait for progress bar to show up, then wait for it to finish
    print("Waiting for progress bar...")
    office.Progress.wait('exists',timeout=6000)
    office.Progress.wait_not('exists',timeout=6000)
    #then finally exit the installer
    print("Exiting the installer...")
    #click_input instead of click so pywinauto doesn't get caught up when the "restart" window opens
    office.CloseButton.wait('exists ready',timeout=6000).click_input()
    #TODO - vv This doesn't work. Need to figure out how to hit "no" on the popup window after closing the main window
    office = app.window(title="Setup")
    if(office.NoButton.exists()):
        office.NoButton.wait('exists ready', timeout=6000).click()
    print("")
        
        
def officePK():
    print("Launching Office 2010 installer for product key...")
    app = Application(backend="uia").start("//***REMOVED***/setup apps/office 2010/setup.exe",timeout=6000)
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
    office.Progress.wait('exists',timeout=6000)
    #then wait for it to finish (not exist)
    office.Progress.wait_not('exists',timeout=6000)
    #click_input instead of click so pywinauto doesn't get caught up when the "restart" window opens
    office.CloseButton.wait('exists ready',timeout=6000).click_input()
    print("Exiting the installer...")
    #TODO - vv This doesn't work. Need to figure out how to hit "no" on the popup window after closing the main window
    office = app.window(title="Setup")
    if(office.NoButton.exists()):
        office.NoButton.wait('exists ready', timeout=6000).click()
    print("")

def sophosInstall():
    print("Launching Sophos installer...")
    app = Application(backend='uia').start("//***REMOVED***/setup apps/Sophos Central Endpoint Installer/SophosSetup.exe")
    #pywinauto has a hard time connecting to sophos, wait 5 seconds to improve likelihood that this works
    time.sleep(5)
    app = Application(backend="uia").connect(title="Sophos Install")
    sophos = app.window()

    print("Clicking the Continue button...")
    sophos.ContinuePane.wait('ready exists',timeout=6000).click_input()
    print("Clicking the Install button...")
    sophos.InstallPane.wait('ready exists',timeout=6000).click_input()
    print("Waiting for progress bar...")
    sophos.InstallationWillTakeAbout10Minutes.wait('exists',timeout=6000)
    sophos.InstallationWillTakeAbout10Minutes.wait_not('exists',timeout=6000)
    #Restart checkbox sometimes does not appear.
    if(sophos.Finish.exists(timeout=6000)):
        if(sophos.RestartMyComputerNow.exists()):
            print("Unchecking restart checkbox, then exiting the installer...")
            #specify coordinates to click the button, because the button is glitched and extends well beyond its actual clickable area (and past the end of the window itself)
            sophos.RestartMyComputerNow.wait('ready exists').click_input(coords=(0,0))
    print("Closing Sophos installer")
    sophos.Finish.wait('ready exists',timeout=6000).click_input()
    print("")
    
    
def zoomInstall():
    print("Launching Zoom installer...")
    app = Application(backend='uia').start("//***REMOVED***/setup apps/zoom/zoominstaller.exe",timeout=6000)
    app = Application(backend='uia').connect(title="Zoom",timeout=6000)
    zoom = app.window(title="Zoom")

    print("Pressing the close button...")
    zoom.Done.wait('ready exists',timeout=6000).click()
    zoom = app.window(title="Zoom Cloud Meetings")
    print("Pressing the other close button...")
    zoom.CloseButton.wait('ready exists',timeout=6000).click()
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
