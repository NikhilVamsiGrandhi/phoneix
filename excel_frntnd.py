from pywinauto.application import Application
import pyautogui as gui
from pywinauto.timings import wait_until
import time


# from datetime import datetime, timedelta
# today = datetime.now() - timedelta(days=1)
# today = today.day
# month = datetime.now().strftime("%B")
# month = month[0:3]
# year = datetime.now().year


# cd desktop/automation/Python/ScrapIT_PTG
# python excel_test.py

applctn = Application(backend='uia').start(r'C:\Program Files\Microsoft Office\root\Office16\EXCEL.exe').connect(title_re=r".*Excel.*",timeout=30)
app = applctn.window(title_re=r".*Excel.*")
app.maximize()


def btnClck(title="", auto_id="", control_type=""):
    try:
        wait_until(
            timeout=60,
            retry_interval=1,
            func=lambda: app.child_window(title=title, auto_id=auto_id, control_type=control_type).exists()
        )
        app.child_window(title=title, auto_id=auto_id, control_type=control_type).click_input()
    except TimeoutError:
        print(f"Element {title} did not appear within the specified timeout.")


btnClck(title="Pinned",control_type="TabItem")
btnClck(title="DailyIntake", control_type="ListItem")

time.sleep(2)

recovery = app.child_window(title="Document Recovery", control_type="Pane").exists()
if recovery=="True":
    app.child_window(title="Document Recovery", control_type="Pane").child_window(title="Close", control_type="Button").click()


# last_sheet = fr"DailyIntake{today}"
# app.child_window(title_re=".*02Mar.*",auto_id="SheetTab", control_type="TabItem").click_input()

text = app.child_window(title_re=r".*S.*", auto_id="S3", control_type="DataItem").window_text()
print(text)

b1 = app.child_window(title_re=r".*B.*", auto_id="B1", control_type="DataItem")
b1.click_input()
gui.press('left')
gui.keyDown('ctrl')
gui.press('a')
gui.press('a')
gui.press('c')
gui.keyUp('ctrl')

btnClck(title="Add Sheet", auto_id="SheetTab", control_type="Button")
b1.click_input()
gui.press('left')
gui.keyDown('ctrl')
gui.press('v')
gui.keyUp('ctrl')

b1.click_input()
gui.typewrite('05-03-2024')
gui.press('enter')

btnClck(title="Sheet1", auto_id="SheetTab", control_type="TabItem").click_input(button='right')

# btnClck(title="Page down", control_type="Button")
# btnClck(title="Page right", control_type="Button")


time.sleep(2)
app.print_control_identifiers()

applctn.kill()