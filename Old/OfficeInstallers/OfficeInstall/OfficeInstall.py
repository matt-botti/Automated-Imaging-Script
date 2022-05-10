#https://stackoverflow.com/a/27923165
#https://pywinauto.readthedocs.io/en/latest/
from pywinauto import Application

app = Application(backend='uia')
#start office setup
print("Starting Office 2010 installer")
app.start("//***REMOVED***/setup apps/office 2010/setup.exe", timeout=300)
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
# vv This doesn't work. Need to figure out how to hit "no" on the popup window after closing the main window
office = app.window(title="Setup")
if(office.NoButton.exists()):
    office.NoButton.wait('exists ready', timeout=300).click()
#TODO - Need to click the checkbox to agree to TOS IF it comes up
#           -It only comes up the first time the installer is run. Even if uninstalled/reinstalled, Office will remember that TOS was agreed to
#TODO - Check if "Restart" prompt comes up after this script runs