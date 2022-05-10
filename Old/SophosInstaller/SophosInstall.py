from pywinauto import Application
import time


#TODO - Figure out how to get pywinauto to properly connect to Sophos. Might need a wait command before connecting?
#Launching Sophos in the PowerShell script also causes a Windows popup, so maybe Application.start() then Application.connect()?
#Update - The method mentioned above works.
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
sophos.Finish.wait('ready exists',timeout=60).click_input()