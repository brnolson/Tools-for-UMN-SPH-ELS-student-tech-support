import pyautogui as pa
import keyboard, time, sys

print(" Press and hold 'esc' to close program\n")
def exit_status():
    if keyboard.is_pressed('esc'):
        exit()

def idle():
    while True:
        print(" idling   ", end='\r')
        pa.drag(20, 0)
        exit_status()
        time.sleep(1.5)
        print(" idling.", end='\r')
        pa.drag(0, 20)
        exit_status()
        time.sleep(1.5)
        print(" idling..", end='\r')
        pa.drag(-20, 0)
        exit_status()
        time.sleep(1.5)
        print(" idling...", end='\r')
        pa.drag(0, -20)
        exit_status()
        time.sleep(1.5)

#-------------------------------------------------------------------------------------------#

if __name__ == "__main__":
    time.sleep(1)
    print(" 3", end='\r')
    time.sleep(1)
    print(" 3 2", end='\r')
    time.sleep(1)
    print(" 3 2 1 ", end='\r')
    time.sleep(1)
    idle()



    
    