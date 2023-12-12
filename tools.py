import pyautogui as pa
import keyboard, time

print(r'    _____            .___.__  __  .__                ___________           .__          ')
print(r'   /  _  \  __ __  __| _/|__|/  |_|__| ____   ____   \__    ___/___   ____ |  |   ______')
print(r'  /  /_\  \|  |  \/ __ | |  \   __\  |/  _ \ /    \    |    | /  _ \ /  _ \|  |  /  ___/')
print(r' /    |    \  |  / /_/ | |  ||  | |  (  <_> )   |  \   |    |(  <_> |  <_> )  |__\___ \ ')
print(r' \____|__  /____/\____ | |__||__| |__|\____/|___|  /   |____| \____/ \____/|____/____  >')
print(r'         \/           \/                         \/                                  \/ ')

# 'a' must delete, 's' must silence, 'f' must insert 1 second of silence, 'm' must create marker, '/' must edit marker

def marker_tool():
    print("\n By: Brenen Olson\n--------------------------------------------\n Press 'n' to insert marker in timeline\n Press '=' to advance slide # and '-' to recede slide #\n Press 'e' to place the final END marker\n Press 'Esc' to leave program\n--------------------------------------------")
    slide = 1
    while True:
        if keyboard.is_pressed('n'):
            if slide == 1:
                pa.press(r'a')
                time.sleep(.05)

                time.sleep(.05)
                pa.press(r'm')
                time.sleep(.05)
                pa.press(r'/')
                if slide < 10:
                    slide_inp = '0' + str(slide)
                else:
                    slide_inp = slide  
                pa.write(f'Slide {slide_inp}')
                pa.press('enter')
                slide += 1
                time.sleep(.05)

                pa.keyDown("shift")
                pa.keyDown("ctrl")
                pa.keyDown("a")
                time.sleep(.05)
                pa.keyUp("shift")
                pa.keyUp("ctrl")
                pa.keyUp("a")

                time.sleep(.1)
                pa.press(r'f')
                time.sleep(.05)
                pa.press('enter')
                time.sleep(.05)

            else:

                pa.press(r'a')
                time.sleep(.05)

                pa.press(r'f')
                time.sleep(.05)
                pa.press('enter')
                time.sleep(.05)
                pa.press('right')
                time.sleep(.05)

                pa.keyDown("shift")
                pa.keyDown("ctrl")
                pa.keyDown("a")
                time.sleep(.05)
                pa.keyUp("shift")
                pa.keyUp("ctrl")
                pa.keyUp("a")

                time.sleep(.05)
                pa.press(r'm')
                time.sleep(.05)
                pa.press(r'/')
                if slide < 10:
                    slide_inp = '0' + str(slide)
                else:
                    slide_inp = slide  
                pa.write(f'Slide {slide_inp}')
                pa.press('enter')
                slide += 1

                pa.keyDown("shift")
                pa.keyDown("ctrl")
                pa.keyDown("a")
                time.sleep(.05)
                pa.keyUp("shift")
                pa.keyUp("ctrl")
                pa.keyUp("a")

                time.sleep(.05)
                pa.press('left')
                time.sleep(.1)
                pa.press(r'f')
                time.sleep(.05)
                pa.press('enter')
                time.sleep(.05)

                pa.keyDown("shift")
                pa.keyDown("ctrl")
                pa.keyDown("a")
                time.sleep(.05)
                pa.keyUp("shift")
                pa.keyUp("ctrl")
                pa.keyUp("a")

        if keyboard.is_pressed('-'):
            if slide > 1:
                slide -= 1
                print(f' Next slide is {slide}')
                time.sleep(.15)


        if keyboard.is_pressed('='):
            if slide > 0:
                slide += 1
                print(f' Next slide is {slide}')
                time.sleep(.15)

        if keyboard.is_pressed('e'):
            pa.press(r'a')
            time.sleep(.05)

            time.sleep(.1)
            pa.press(r'f')
            time.sleep(.05)
            pa.press('enter')

            time.sleep(.05)
            pa.press(r'end')
            time.sleep(.05)
            pa.press(r'm')
            time.sleep(.05)
            pa.press(r'/')

            pa.write('END')
            pa.press('enter')
            break

        if keyboard.is_pressed('esc'):
            break

#-------------------------------------------------------------------------------------------#

if __name__ == "__main__":
    marker_tool()



    
    