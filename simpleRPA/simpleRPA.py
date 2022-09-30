import pyautogui
import pyperclip
import xlrd
import time

# This is a simple Robotic Process Automation program.
# Based on the current screenshots and data in the sheet, it will click the like button to all the comments under a YouTube video.
# It can also do other tasks if you change the screenshots and data in the sheet.


# Check the data first
# actionType.value: 1.0: single left click; 2.0: double left click; 3.0 single right click;
# # 4.0: type words; 5.0: wait; 6.0: scroll
# actionType.ctype: 0: empty; 1: string; 2: float;
def dataCheck(sheet1):
    isValid = True
    if sheet1.nrows<2:
        print("No Data")
        isValid = False
    # Start from the second row
    i = 1
    while i < sheet1.nrows:
        # First col, check the action type
        actionType = sheet1.row(i)[0]
        # Action can only be an float and from 1.0 to 6.0
        if actionType.ctype != 2 or (actionType.value != 1.0 and actionType.value != 2.0 and actionType.value != 3.0
        and actionType.value != 4.0 and actionType.value != 5.0 and actionType.value != 6.0):
            print("Data in Row " + str(i+1) + ", " + "Col 2 is wrong")
            isValid = False
        # Second col, check the content
        actionContent = sheet1.row(i)[1]
        # For the click action, ctype must be a string
        if actionType.value == 1.0 or actionType.value == 2.0 or actionType.value == 3.0:
            if actionContent.ctype != 1:
                print("Data in Row " + str(i+1) + ", " + "Col 2 is wrong")
                isValid = False
        # For the input words action, ctype cannot be empty
        if actionType.value == 4.0:
            if actionContent.ctype == 0:
                print("Data in Row " + str(i+1) + ", " + "Col 2 is wrong")
                isValid = False
        # For the wait action, ctype must be an float
        if actionType.value == 5.0:
            if actionContent.ctype != 2:
                print("Data in Row " + str(i+1) + ", " + "Col 2 is wrong")
                isValid = False
        # For the scroll action, ctype must be an float
        if actionType.value == 6.0:
            if actionContent.ctype != 2:
                print("Data in Row " + str(i+1) + ", " + "Col 2 is wrong")
                isValid = False
        i += 1
    return isValid


# Mouse Actions
# clickTimes: single or double click;
# lOrR: left or right click;
# img: the image that mouse click on;
# repeat: the number of executions
def mouseClick(clickTimes,lOrR,img,repeat):
    # Execute once
    if repeat == 1:
        while True:
            location = pyautogui.locateCenterOnScreen(img,confidence=0.9)
            if location is not None:
                pyautogui.click(location.x,location.y,clicks=clickTimes,interval=0.2,button=lOrR)
                break
            print("Unable to find the image, retry in 0.1 second")
            time.sleep(0.1)
    # Execute infinite times
    elif repeat == -1:
        while True:
            location=pyautogui.locateCenterOnScreen(img,confidence=0.9)
            if location is not None:
                pyautogui.click(location.x,location.y,clicks=clickTimes,interval=0.2,button=lOrR)
            time.sleep(0.1)
    # Execute certain times
    elif repeat > 1:
        i = 1
        while i < repeat + 1:
            location=pyautogui.locateCenterOnScreen(img,confidence=0.9)
            if location is not None:
                pyautogui.click(location.x,location.y,clicks=clickTimes,interval=0.2,button=lOrR)
                print("Repeat the action")
                i += 1
            time.sleep(0.1)



def mainFunction(img):
    i = 1
    while i < sheet1.nrows:
        # Get the action type
        actionType = sheet1.row(i)[0]
        # Left single click
        if actionType.value == 1:
            # Get the image name
            img = sheet1.row(i)[1].value
            # Repeat numbers
            repeat = 1
            # Check if it repeats
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                repeat = sheet1.row(i)[2].value
            mouseClick(1,"left",img,repeat)
            print("left single click",img)
        # Left double click
        elif actionType.value == 2:
            img = sheet1.row(i)[1].value
            repeat = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                repeat = sheet1.row(i)[2].value
            mouseClick(2,"left",img,repeat)
            print("left double click",img)
        # Right single click
        elif actionType.value == 3.0:
            img = sheet1.row(i)[1].value
            repeat = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                repeat = sheet1.row(i)[2].value
            mouseClick(1,"right",img,repeat)
            print("right click",img)
        # Input words
        elif actionType.value == 4.0:
            inputWords = sheet1.row(i)[1].value
            pyperclip.copy(inputWords)
            pyautogui.hotkey('ctrl','v')
            time.sleep(0.5)
            print("input:",inputWords)
        # Wait
        elif actionType.value == 5.0:
            waitTime = sheet1.row(i)[1].value
            time.sleep(waitTime)
            print("wait",waitTime,"second")
        # Scroll
        elif actionType.value == 6.0:
            distance = sheet1.row(i)[1].value
            pyautogui.scroll(int(distance))
            print("Scrolling",int(distance),"distance")
        i += 1


if __name__ == '__main__':
    file = 'actions.xls'
    # Open the file
    wb = xlrd.open_workbook(filename = file)
    # Get the sheet
    sheet1 = wb.sheet_by_index(0)
    print('Welcome to the simple RPA')
    # Check the data
    isValid = dataCheck(sheet1)
    if isValid:
        key = input('Press 1 to start: \n')
        if key == '1':
            mainFunction(sheet1)
    else:
        print('Something wrong in the data')
