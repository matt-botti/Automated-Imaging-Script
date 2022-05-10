from pywinauto import Application

app = Application(backend='uia').connect(title="Add a device")
printer = app.window()

printer.FollowYouPrinterOnEquitracButton.wait('ready exists', timeout=60).click()
printer.NextButton.wait('ready exists', timeout=30).click()

printer = app.window(title="Add Printer")
printer.FinishButton.wait('ready exists',timeout=60).click()

