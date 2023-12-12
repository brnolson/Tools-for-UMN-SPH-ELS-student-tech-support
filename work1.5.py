import pyautogui as pa
import time, os

#------------------------------------------------------------------------------------#
# Must have Powerpoint and Audition open and on their main pages to function properly
# Powerpoint must be in slide view
# Audition must be in wave view and all audio must be merged
# Markers in audition must be in the top left panel
# Both must be open on the RIGHT monitor (display 1)
#------------------------------------------------------------------------------------#

def file_structure():
    exists = os.listdir(r"C:\Users\ols00175\Desktop")
    if "Development" not in exists:
        os.mkdir("Development")
        os.mkdir("Production")
        os.mkdir("Source")
    else:
        print("Folders already exist")

#-------------------------------------------------------------------------------------------#

def pdf_setup(type):
    windows = pa.getAllWindows()

    powerpoint_window = None
    for window in windows:
        if "PowerPoint" in window.title:
            powerpoint_window = window
            break

    if powerpoint_window is not None:
        powerpoint_window.maximize()
        time.sleep(.5)
        pa.click(494,17, duration=.25)
        time.sleep(.25)
        pa.click(31,42, duration=.25)
        time.sleep(.1)
        pa.click(49,244)
        time.sleep(.1)
        pa.doubleClick(244,205)
        time.sleep(.2)

# from top left
        location = pa.getActiveWindow().left+125, pa.getActiveWindow().top+50
        pa.moveTo(location)
        pa.click()
        time.sleep(.1)
        pa.write(r"C:\Users\ols00175\Desktop")
        time.sleep(.1)
        pa.press('enter')

# from bottom left
        location = pa.getActiveWindow().left+150, pa.getActiveWindow().bottom-110
        pa.click(location)
        if type == 'png':
            return
        if pa.getActiveWindow().bottom < 760:
            pa.move(0, 62)
            pa.click()
        else:
            pa.move(0, -385)
            pa.click()
        location = pa.getActiveWindow().left+150, pa.getActiveWindow().bottom-120
        pa.click(location)
        time.sleep(.1)

# to ensure drop down menu goes down
        if pa.getActiveWindow().bottom > 720:
            location = pa.getActiveWindow().top+10, pa.getActiveWindow().left+30
            pa.moveTo(location)
            pa.mouseDown()
            pa.moveTo(760, 300)
            pa.mouseUp()

        location = pa.getActiveWindow().left+132, pa.getActiveWindow().bottom-245
        pa.click(location)

    else:
        print("PowerPoint window not found")
        return False

def create_slides():
    pdf_setup('png')
    if pa.getActiveWindow().bottom < 760:
        pa.move(0, 305)
        pa.click()
    else:
        pa.move(0, -145)
        pa.click()
    location = pa.getActiveWindow().right-140, pa.getActiveWindow().bottom-140
    pa.moveTo(location)
    time.sleep(.1)
    pa.click()
    time.sleep(.5)
    pa.click()
    for i in range(5):
        pa.press('backspace')
    pa.write("FinalSlides")
    location = pa.getActiveWindow().right-150, pa.getActiveWindow().bottom-30
    pa.click(location)
    time.sleep(.2)
    location = pa.getActiveWindow().left+100, pa.getActiveWindow().bottom-30
    pa.click(location)
    time.sleep(4)
    location = pa.getActiveWindow().right-10, pa.getActiveWindow().top+10
    pa.click(location)
    time.sleep(.3)

def create_transcript():
    status = pdf_setup("pdf")
    if status == False:
        print("Operation could not be performed")
        return
    pa.move(0, 50, duration=.1)
    pa.click()
    location = pa.getActiveWindow().right-140, pa.getActiveWindow().bottom-30
    pa.moveTo(location)
    pa.click()
    location = pa.getActiveWindow().right-140, pa.getActiveWindow().bottom-215
    pa.moveTo(location)
    pa.click()
    time.sleep(.5)
    pa.click()
    for i in range(5):
        pa.press('backspace')
    pa.write("Transcript")
    location = pa.getActiveWindow().right-150, pa.getActiveWindow().bottom-30
    pa.click(location)

def create_handouts():
    status = pdf_setup("pdf")
    if status == False:
        print("Operation could not be performed")
        return
    pa.move(0, 35, duration=.1)
    pa.click()
    location = pa.getActiveWindow().right-132, pa.getActiveWindow().bottom-245
    pa.click(location, duration=.15)
    pa.move(0, 50)
    pa.click()
    location = pa.getActiveWindow().right-140, pa.getActiveWindow().bottom-30
    pa.moveTo(location)
    pa.click()
    location = pa.getActiveWindow().right-140, pa.getActiveWindow().bottom-215
    pa.moveTo(location)
    pa.click()
    time.sleep(.5)
    pa.click()
    for i in range(5):
        pa.press('backspace')
    pa.write("Handouts")
    location = pa.getActiveWindow().right-150, pa.getActiveWindow().bottom-30
    pa.click(location)

def chill():
    window = pa.getActiveWindowTitle()
    while "PowerPoint" in window:
        window = pa.getActiveWindowTitle()
    window = pa.getActiveWindowTitle()
    while window == "Publishing...":
        window = pa.getActiveWindowTitle()
    time.sleep(.1)

def minimize_window():
    time.sleep(.5)
    pa.click(1298,17, duration=.5)
    location = pa.getActiveWindow().right-130, pa.getActiveWindow().top+20
    pa.click(location)

def powerpoint():
    print("    Powerpoint")

    path = os.listdir(r"C:\Users\ols00175\Desktop\Development")
    exists = False
    for item in path:
        if "FinalSlides" in item:
            exists = True
            print("Slides already created")
    print("\t.")
    if exists != True:
        print("Creating slides")
        create_slides()

    path = os.listdir(r"C:\Users\ols00175\Desktop\Production")
    exists = False
    for item in path:
        if "Handouts" in item:
            exists = True
            print("Handouts already created")
    print("\t.")
    if exists != True:
        print("Creating handouts")
        create_handouts()
    print("\t.")

    path = os.listdir(r"C:\Users\ols00175\Desktop\Production")
    exists = False
    for item in path:
        if "Transcript" in item:
            exists = True
            print("Transcript already created")
    if exists != True:
        chill()
        print("Creating transcript")
        create_transcript()
        chill()

    print("\t.")
    minimize_window()

#-------------------------------------------------------------------------------------------#

def audio_export():
    windows = pa.getAllWindows()

    audition_window = None
    for window in windows:
        if "Audition" in window.title:
            audition_window = window
            break

    if audition_window is not None:
        audition_window.maximize()
        time.sleep(.5)
        pa.click(1298,17, duration=.25)
        pa.click(100,146, duration=.1)
        pa.hotkey("ctrl", "a")
        pa.click(129,83, duration=.1)
        print("\t.")
        location = pa.getActiveWindow().right-75, pa.getActiveWindow().top+170
        pa.click(location)
        time.sleep(.15)
        location = pa.getActiveWindow().left+125, pa.getActiveWindow().top+50
        pa.moveTo(location)
        pa.click()
        time.sleep(.1)
        pa.write(r"C:\Users\ols00175\Desktop\Development")
        time.sleep(.1)
        pa.press('enter')
        print("\t.")
        location = pa.getActiveWindow().left+200, pa.getActiveWindow().top+150
        pa.click(location, duration=.15)
        time.sleep(.5)
        location = pa.getActiveWindow().right-50, pa.getActiveWindow().bottom-60
        pa.moveTo(location)
        pa.click()
        time.sleep(.5)
        pa.click()
        for i in range(6):
            pa.press('backspace')
        pa.write("Audio")
        pa.hotkey("ctrl", "a")
        pa.hotkey("ctrl", "c")
        print("\t.")

        pa.move(-50,-250)
        pa.rightClick()
        pa.move(20,215)
        time.sleep(.75)
        pa.move(300, 0)
        pa.click()
        time.sleep(.5)
        pa.hotkey("ctrl", "v")
        pa.press('enter')
        print("\t.")

        location = pa.getActiveWindow().right-175, pa.getActiveWindow().bottom-35
        pa.click(location, duration=.15)
        time.sleep(.5)
        location = pa.getActiveWindow().right-150, pa.getActiveWindow().bottom-30
        pa.click(location, duration=.15)
        minimize_window()

    else:
        print("Audition window not found")
        return False

def audition():
    print("    Audition")
    audio_export()

#-------------------------------------------------------------------------------------------#

if __name__ == "__main__":
    file_structure()
    choice = input('Enter "p" for powerpoint process, "a" for audition process, or "b" for both. ')
    while choice != "b" and choice != "a" and choice != "p":
        print("Invalid input...")
        choice = input('Enter "p" for powerpoint process, "a" for audition process, or "b" for both. ')
    if choice == "p":
        powerpoint()
    elif choice == "a":
        audition()
    elif choice == "b":
        powerpoint()
        audition()
    print("done!")



    
    