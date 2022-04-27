import datetime, win32gui, win32con, pyautogui, cv2, pytesseract, json, logging, os
from cv2 import split
from socket import timeout
from pytesseract import Output
from pywinauto import application, findwindows
from genericpath import exists
from PIL import Image

_defaultConfigFileName = "configCloseAnyWindow.json"
_defaultConfigContent = {
	"Windows" : [{
		"WindowName":"C22(1)/747 Abbruch",
		"Type":"OCR",
		"OcrButtonName": "Beenden",
        "OcrOffset": "30",
        "TimeOutInSeconds": "10"
	},{
		"WindowName":"Calculator",
		"Type":"FORCE",
        "TimeOutInSeconds": "10"
	},{
		"WindowName":"Calculator",
		"Type":"HotKey",
        "HotKeyHoldButton":"",
        "HotKeyPressButton":"1|1|3|4",
        "TimeOutInSeconds": "10",
        "DocLink": "https://pyautogui.readthedocs.io/en/latest/keyboard.html"
	}
    ,{
		"WindowName":"Calculator",
		"Type":"HotKey",
        "HotKeyHoldButton":"Alt",
        "HotKeyPressButton":"f4",
        "TimeOutInSeconds": "10",
        "DocLink": "https://pyautogui.readthedocs.io/en/latest/keyboard.html"
	}],
    "DelayBeforeStartInSeconds": "3",
	"AfterRunDeleteJPG": "False"
}

_screenshot = "screenshot.jpg"
_screenshotWhite = "screenshotwhite.jpg"
_currentWindow = "currentWindow.jpg"

#_tesseractPath = r"C:\Users\z004e9mc\AppData\Local\Programs\Tesseract-OCR\tesseract"
_tesseractPath = r'Tesseract-OCR\tesseract'

logging.basicConfig(filename='log.txt', encoding='utf-8', level=logging.DEBUG, format='%(asctime)s %(levelname)-8s %(message)s')

def CloseAnyWindowByForce(window_name, timeout):
    closed = False
    try:
        handles = findwindows.find_windows(title=window_name)
    except Exception as exp:
        logging.error('Window Name [' + window_name + '] Type [Force] Error when try find window with title [' + str(exp) + ']')
        pass #Just do nothing if the pop-up dialog was not found
    else: #The window was found, so click the button
        for handle in handles: 
            try: 
                app = application.Application()
                app.Connect(handle=handle)
                currentWindow = app[window_name]
                if(currentWindow.exists):
                    win32gui.SetForegroundWindow(handle)
                    win32gui.SendMessage(handle, win32con.WM_CLOSE,0,0)
                    if(len(findwindows.find_windows(title=window_name)) == 0):
                        logging.info('Window Name [' + window_name + '] Type [Force]  Closed' )
                        closed = True
                        break                    
            except Exception as exp:
                logging.error('Window Name [' + window_name + '] Type [Force] Error when try close window with title [' + str(exp) + ']')
                pass #Just do nothing if cant close window

    pyautogui.sleep(1) #should help reduce cpu load a little for this thread

    if(datetime.datetime.now().time() < timeout.time()):
        if(not closed):
            logging.warning('Window Name [' + window_name + '] Type [Force] NOT Closed')
            CloseAnyWindowByForce(window_name, timeout)
    else:
        logging.warning('Window Name [' + window_name + '] Type [Force] Timeout')

def CloseAnyWindowByOCR(window_name, button_name, offset, timeout):
    closed = False           
    try:
        handles = findwindows.find_windows(title=window_name)
    except Exception as exp:
        logging.error('Window Name [' + window_name + '] Button Name [' + button_name+ '] Type [OCR]  Error when try find window with title [' + str(exp) + ']')
        pass #Just do nothing if the pop-up dialog was not found
    else: #The window was found, so click the button
        for handle in handles:
            try:            
                #Get postion fo Current Window
                win32gui.SetForegroundWindow(handle)
                xCW, yCW, xCW1, yCW1 = win32gui.GetClientRect(handle)
                xCW, yCW = win32gui.ClientToScreen(handle, (xCW, yCW))
                xCW1, yCW1 = win32gui.ClientToScreen(handle, (xCW1 - xCW, yCW1 - yCW))

                #Full screenshot
                imageScreenshot = pyautogui.screenshot()
                imageScreenshot.save(_screenshot)
                imageScreenshot = cv2.imread(_screenshot)

                #Full screenshot as white
                imageScreeshotWhite = imageScreenshot.copy()
                imageScreeshotWhite.fill(255)
                cv2.imwrite(_screenshotWhite, imageScreeshotWhite)

                #Screenshot only for current window
                imageCurrentWindow = pyautogui.screenshot(region=(xCW, yCW- offset, xCW1, yCW1+ offset))
                imageCurrentWindow.save(_currentWindow)

                #Merge images and save as currentWindow
                imageScreeshotWhite = Image.open(_screenshotWhite)
                imageCurrentWindow = Image.open(_currentWindow)
                imageScreeshotWhite.paste(imageCurrentWindow, (xCW, yCW))
                imageScreeshotWhite.convert(imageScreeshotWhite.mode).save(_currentWindow)

                #Modify currentwindow to help find words
                image = cv2.imread(_currentWindow)
                imageGray  = cv2.cvtColor(image , cv2.COLOR_BGR2GRAY)
                imageThreshold  = cv2.threshold(imageGray, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)[1]
                imageThreshold = cv2.bilateralFilter(imageThreshold, 11, 20, 20)
                
                cv2.imwrite(_screenshotWhite, imageThreshold)

                #Extraxt words from currentWindow
                customConfig = r'--oem 3 --psm 6' #3 Default, based on what is available. 6 Assume a single uniform block of text.
                pytesseract.pytesseract.tesseract_cmd = (_tesseractPath)
                data = pytesseract.image_to_data(imageThreshold, output_type=Output.DICT, config=customConfig)
                logging.info('Window Name [' + window_name + '] Button Name [' + button_name+ '] Type [OCR] OCR words found:' )
                logging.info('OCR words [' + str(data['text']) + ']')

                #Find a button to click
                boxes = len(data['level'])
                for i in range(boxes):
                    (x, y, width, height) = (data['left'][i], data['top'][i], data['width'][i], data['height'][i])
                    if (data['text'][i]).upper().strip() == button_name.upper().strip():
                        click(str(x+(width/2)) + '|' + str(y+(height/2)- offset))
                        if(len(findwindows.find_windows(title=window_name)) == 0):
                            logging.info('Window Name [' + window_name + '] Button Name [' + button_name+ '] Type [OCR] Closed' )
                            closed = True
                            break
                
            except Exception as exp:
                logging.error('Window Name [' + window_name + '] Button Name [' + button_name+ '] Type [OCR]  Error when try find button [' + str(exp) + ']')
                pass #Just do nothing if cant close window 
    
    pyautogui.sleep(1) #should help reduce cpu load a little for this thread

    if(datetime.datetime.now().time() < timeout.time()):
        if(not closed):
            logging.warning('Window Name [' + window_name + '] Button Name [' + button_name+ '] Type [OCR] NOT Closed' )
            CloseAnyWindowByOCR(window_name, button_name, offset, timeout)
    else:
         logging.warning('Window Name [' + window_name + '] Button Name [' + button_name+ '] Type [OCR] Timeout' )                    

 
def CloseAnyWindowByHotkey(window_name, button_hold, button_press, timeout):
    closed = False
    try:
        handles = findwindows.find_windows(title=window_name)
    except Exception as exp:
        logging.error('Window Name [' + window_name + '] Type [Hotkey] Error when try find window with title [' + str(exp) + ']')
        pass #Just do nothing if the pop-up dialog was not found
    else: #The window was found, so click the button
        for handle in handles: 
            try: 
                app = application.Application()
                app.Connect(handle=handle)
                currentWindow = app[window_name]
                if(currentWindow.exists):
                    #Set focus when click on window
                    win32gui.SetForegroundWindow(handle)
                    x, y, x1, y1 = win32gui.GetClientRect(handle)
                    x, y = win32gui.ClientToScreen(handle, (x, y))
                    click(str(x) + '|' + str(y))

                    if(len(button_hold) > 0):
                        with pyautogui.hold(button_hold):
                            pyautogui.press(str(button_press).split("|"))
                    else:
                        pyautogui.press(str(button_press).split("|"))

                    if(len(findwindows.find_windows(title=window_name)) == 0):
                        logging.info('Window Name [' + window_name + '] Type [Hotkey]  Closed' )
                        closed = True
                        break                    
            except Exception as exp:
                logging.error('Window Name [' + window_name + '] Type [Hotkey] Error when try close window with title [' + str(exp) + ']')
                pass #Just do nothing if cant close window

    pyautogui.sleep(1) #should help reduce cpu load a little for this thread

    if(datetime.datetime.now().time() < timeout.time()):
        if(not closed):
            logging.warning('Window Name [' + window_name + '] Type [Hotkey] NOT Closed')
            CloseAnyWindowByHotkey(window_name, button_hold, button_press, timeout)
    else:
        logging.warning('Window Name [' + window_name + '] Type [Hotkey] Timeout')   

def click(coordenada):
        x = int(float(coordenada.split('|')[0]))
        y = int(float(coordenada.split('|')[1]))
        pyautogui.moveTo(x,y)
        pyautogui.click(x,y)

try: 
    if(not exists(_defaultConfigFileName)): #Just create a config file if not exist
        with open(_defaultConfigFileName, 'w', newline='') as file:
                file.write(json.dumps(_defaultConfigContent))

    jsonFile = open(_defaultConfigFileName)
    jsonData = json.load(jsonFile)

    pyautogui.sleep(int(jsonData['DelayBeforeStartInSeconds']))
    for windowData in jsonData['Windows']:        
        match str(windowData['Type']).upper().strip():
            case "OCR":
                timeout = datetime.datetime.now() + datetime.timedelta(seconds=int(windowData['TimeOutInSeconds']))
                CloseAnyWindowByOCR(windowData['WindowName'], windowData['OcrButtonName'], int(windowData['OcrOffset']), timeout)
            case 'FORCE':
                timeout = datetime.datetime.now() + datetime.timedelta(seconds=int(windowData['TimeOutInSeconds']))
                CloseAnyWindowByForce(windowData['WindowName'], timeout)
            case 'HOTKEY':
                timeout = datetime.datetime.now() + datetime.timedelta(seconds=int(windowData['TimeOutInSeconds']))
                CloseAnyWindowByHotkey(windowData['WindowName'], windowData['HotKeyHoldButton'], windowData['HotKeyPressButton'], timeout)
            case _:
                logging.warning('Type [' + str(windowData['Type']) + '] not exit' )                    
    jsonFile.close

    # remove .jpg files
    if(str(jsonData["AfterRunDeleteJPG"]).upper().strip() == "TRUE" ):
        files_in_directory = os.listdir('./')
        filtered_files = [file for file in files_in_directory if file.endswith(".jpg")]
        for file in filtered_files:
            path_to_file = os.path.join('./', file)
            os.remove(path_to_file)

except Exception as exp:
    logging.error(str(exp))