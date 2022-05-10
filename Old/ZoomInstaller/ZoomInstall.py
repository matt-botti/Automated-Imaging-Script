from pywinauto import Application

print("Launching Zoom installer...")
app = Application(backend='uia').start("//***REMOVED***/setup apps/zoom/zoominstaller.exe",timeout=300)
app = Application(backend='uia').connect(title="Zoom",timeout=600)
zoom = app.window(title="Zoom")

print("Pressing the close button...")
zoom.Done.wait('ready exists',timeout=300).click()
zoom = app.window(title="Zoom Cloud Meetings")
print("Pressing the other close button...")
zoom.CloseButton.wait('ready exists',timeout=60).click()
