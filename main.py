import pyautogui as gui
from pywinauto.timings import wait_until
from pywinauto.application import Application
# from excel import new_excel_sheet
import pytesseract
from PIL import ImageGrab
import time
from datetime import datetime, timedelta


excel_file_path = r"C:\Users\KSPL\Desktop\automation\Python\ScrapIT_PTG\excel_files\DailyIntake_test.xlsx"
pdf_file_path =""
today = datetime.now().day
month = datetime.now().strftime("%B")
tommorrow = datetime.now() + timedelta(days=1)
tommorrow = datetime.now().day

def pClick(x,y,sleepTIme=1):
    gui.moveTo(x,y)
    gui.click()
    time.sleep(sleepTIme)
    
def handle_date_select(duration):
    if tommorrow !=1:
        #from date
        if duration == "day":        
            pClick(585,255)
            gui.press('left')
            gui.press('enter')
        else:
            pClick(585,255)
            for _ in range(today):
                gui.press('left')
            gui.press('enter')

        #to date
        pClick(923,255)
        gui.press('left')
        gui.press('enter')
    else:
        #from date
        if duration == "day":        
            pClick(585,255)
            pClick(585,255)
            gui.press('right')
            gui.press('left')
            for _ in range(today):
                gui.press('right')
            gui.press('enter')
        else:
            pClick(585,255)
            pClick(585,255)
            gui.press('right')
            gui.press('left')
            gui.press('enter')
        #to date
        pClick(923,255)
        pClick(923,255)
        gui.press('right')
        gui.press('left')
        gui.press('enter')

    

def pdf_download(duration):    
    global pdf_file_name
    applctn = Application(backend='uia').start(r'C:\Program Files (x86)\SCRAPIT GUI\SCRAPIT GUI.exe').connect(title='Current version is v1.23.31',timeout=60)
    app = applctn.window(title='Current version is v1.23.31')

    
    pdf_file_path = r"C:\Users\KSPL\Desktop\automation\Python\ScrapIT_PTG\pdfs"
    pdf_file_name = rf"Daily_Product_Purchase_By_Date[{today} {month[0:3]}]" if duration == "day" else rf"MTD_Product_Purchase_By_Date[{today}_{month[0:3]}]"
    
    time.sleep(2)
    
    #Login
    gui.typewrite("Chirag")
    gui.press('tab')
    gui.typewrite("Phoenix")
    time.sleep(1)
    gui.press('enter')

    time.sleep(3)

    pClick(974,612)

    #processing
    pClick(975,165)
    # scale Ticket reporting
    pClick(1050,685)
    # Scale ticket analysis reporting
    pClick(480,350)
    # Product by purchase date
    pClick(480,400)
    
    handle_date_select(duration)
    
    # allocation checkbox
    pClick(503,326)
    #gst checkbox
    pClick(1366,623)
    #submit
    pClick(1545,173,5)

    #pdf on screen
    while True:
        screen_area = ImageGrab.grab(bbox=(221,116,1695,990))
        text = pytesseract.image_to_string(screen_area)
        if "PDF On Screen" in text:
            time.sleep(8)
            gui.moveTo(430,218)
            gui.doubleClick()
            break
        else:
            time.sleep(2)

    #pdf view
    while True:
        screen_area = ImageGrab.grab(bbox=(221,116,1695,990))
        text = pytesseract.image_to_string(screen_area)
        assert "Create VIN Export" not in text, "No data issue"
        assert "FaultEvent" not in text, "Issue at application"
        if "Click Here To View PDF" in text:
            pClick(987,504)
            break
        else:
            time.sleep(3)
            

    partial_title = "Google Chrome"
    app = Application(backend="uia").connect(title_re=fr".*{partial_title}.*",timeout=30)
    chrome = app.window(title_re=fr".*{partial_title}.*")

    #chorme pdf view
    while True:
        screen_area = ImageGrab.grab(bbox=(10,10,1900,1000))
        text = pytesseract.image_to_string(screen_area)
        if "Phoenix" in text or "TRANSPORT" in text:
            pClick(1767,177)
            break
        else:
            time.sleep(3)

    #chorme download tab
    # chrome.child_window(title="Download", auto_id="download", control_type="Button").click()    
    while True:
        screen_area = ImageGrab.grab(bbox=(0,0,150,150))
        text = pytesseract.image_to_string(screen_area)
        if "Save As" in text:
            break
        else:
            time.sleep(3)
            
    try:
        wait_until(
            timeout=60,
            retry_interval=1,
            func=lambda: chrome.child_window(title="Previous Locations", control_type="Button").exists()
        )
        chrome.child_window(title="Previous Locations", control_type="Button").click_input()
    except TimeoutError:
        print(f"Element did not appear within the specified timeout.")    
    gui.typewrite(pdf_file_path)
    gui.press('enter')
    
    chrome.child_window(title="File name:", auto_id="1001", control_type="Edit").type_keys(rf"{pdf_file_name}")
    time.sleep(1)
    gui.press('enter')
    
    # pClick(710,686)
    
    try:
        wait_until(
            timeout=60,
            retry_interval=1,
            func=lambda: chrome.child_window(title_re=".*Keep.*", control_type="Button").exists()
        )
        chrome.child_window(title_re=".*Keep.*", control_type="Button").click_input()
    except TimeoutError:
        print(f"Element did not appear within the specified timeout.")
    
    time.sleep(2)
    
    try:
        wait_until(
            timeout=60,
            retry_interval=1,
            func=lambda: chrome.child_window(title_re=".*pdf.*", control_type="TabItem").child_window(title="Close", control_type="Button").exists()
        )
        chrome.child_window(title_re=".*pdf.*", control_type="TabItem").child_window(title="Close", control_type="Button").click_input()
    except TimeoutError:
        print(f"Element did not appear within the specified timeout.")
    applctn.kill()
    
# pdf_download("month")
pdf_download("day")
# new_excel_sheet(excel_file_path,pdf_file_path,"day")

# pdf_download("month")
# new_excel_sheet(excel_file_path,pdf_file_path,"month")