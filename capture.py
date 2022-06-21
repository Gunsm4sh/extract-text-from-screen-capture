import win32api
import win32con

from PIL import ImageGrab
import pytesseract
import csv


header = ['time', 'wind speed']


f = open('C:\pzthon capture\data2.csv', 'w', newline='',encoding='UTF8')

writer = csv.writer(f)
writer.writerow(header)
custom_config = "outputbase digits"

#win32api.SetCursorPos((650,460))
# win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,750,700,0,0)
# win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,750,700,0,0)
for x in range(4380):
    bbox = (20, 148, 90, 162)
    imdate = ImageGrab.grab(bbox)
    bbox = (500, 148, 530, 162)
    im = ImageGrab.grab(bbox)
    left = pytesseract.image_to_string(imdate)
    right = pytesseract.image_to_string(im,config=custom_config)
    print(left.replace('\n', ''))
    try:
        temp = float(right.replace('\n', ''))
    except:
        print('skipped')
        print('-------------')
        data = [left.replace('\n', ''),'skipped']
        writer.writerow(data)
    else:
        if(right.find('.') == -1):
            temp  = temp/10
        if(temp ==0):
            data = [left.replace('\n', ''),'skipped']
            print('skipped')
        else:
            data = [left.replace('\n', ''),temp]
            print(temp)
        print('-------------')
        writer.writerow(data)
    win32api.SetCursorPos((650,460))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,650,460,0,0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,650,460,0,0)

print("done")
