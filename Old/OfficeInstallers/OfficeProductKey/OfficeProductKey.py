#https://stackoverflow.com/a/27923165
#https://pywinauto.readthedocs.io/en/latest/
from pywinauto import Application

print("Starting Office 2010 installer...")
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
# vv This doesn't work. Need to figure out how to hit "no" on the popup window after closing the main window
office = app.window(title="Setup")
if(office.NoButton.exists()):
    office.NoButton.wait('exists ready', timeout=300).click()
#TODO - Check if "Restart" prompt comes up after this script runs