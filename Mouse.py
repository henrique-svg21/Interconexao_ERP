#just to find out the coordinates
import pyautogui, sys
print('Press Ctrl-C to quit.')
try:
    while True:
        x, y = pyautogui.position()
        positionStr = 'X: ' + str(x).rjust(4) + ' Y: ' + str(y).rjust(4)
        print(positionStr, end='')
        print('\b' * len(positionStr), end='', flush=True)
except KeyboardInterrupt:
    print('\n')

#MOUSE: 381,156

#CONFIRMING EMAIL
#x = 390
#y = 640

#YES: 390,372
#NO: 390, 587