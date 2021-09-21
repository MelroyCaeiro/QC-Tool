# QR Code & OCR scanner v2.0
# Works best with a tripod; keep both QR code and label in frame
#
# Press 'Q' to scan QR code and OCR text in one frame
# Press 'P' to correct the beacon name if OCR fails (hotkey can be changed)
# Press 'N' to save the workbook ONLY AT THE END OF THE BATCH (script will have to be restarted after this; hotkey can be changed)
#
# Download an IP webcam app to use your phone camera (IP Webcam: https://play.google.com/store/apps/details?id=com.pas.webcam&hl=en&gl=US)
# Set video resolution to 640x480, and orientation to 'portrait'
# Edit 'url' based on the IP address shown
#
# Melroy Caeiro, 210921-0934

import cv2
import numpy as np
from pyzbar.pyzbar import decode
import easyocr
import cv2
import xlsxwriter
from time import time, ctime
import keyboard
import beepy as beep
#import warnings
#warnings.filterwarnings("ignore")

t = time()

workbook = xlsxwriter.Workbook('Scan-{}.xlsx'.format(t))  # Excel workbook name; {} will autofill with the number of seconds since epoch
worksheet = workbook.add_worksheet()
worksheet.write('A1', 'QR Code')
worksheet.write('B1', 'Beacon Name')

keyboard.add_hotkey("n", lambda: workbook.close())      # Hotkey to save the Excel workbook; press only when you're actually done with all scans (need to restart script once pressed)

url = 'http://192.168.x.x:8080/video'      # Phone camera IP address
cap = cv2.VideoCapture(url)     # '0' for PC webcam, 'url' for IP webcam

currentData = b'bruh'
substring = 'B'
text = '?'

workStatus = True
ocrStatus = False
excelStatus = False
editStatus = False
counter = 1
editCounter = 1

keyboard.add_hotkey("p", lambda : worksheet.write('B{}'.format(editCounter), input("Correct Beacon Name: ")))     # Hotkey to edit the current beacon name entry, if OCR fails

try:
    while (workStatus == True):
        counter += 1

        success, img = cap.read()
        cv2.imshow('QR Code', img)
        key = cv2.waitKey(33)

        if key == 'q':
            break
        elif key != -1:
            # img2 = cv2.imshow('OCR', img)
            # cv2.waitKey(1)
            for barcode in decode(img):
                myData = barcode.data.decode('UTF-8')

                if (myData == b''):
                    print("QR Code not detected!")
                    counter -= 1
                    break
                elif (myData != currentData):
                    primaryColor = (255, 0, 255)
                    print("QR: ", myData)
                    currentData = myData
                    worksheet.write('A{}'.format(counter), myData)

                    reader = easyocr.Reader(['en'], gpu=True)
                    result = reader.readtext(img)
                    for detection in result:
                        text = detection[1]
                    print("Beacon: ", text, "\n")
                    worksheet.write('B{}'.format(counter), text)
                    editCounter = counter
                    ocrStatus = False
                    # cv2.destroyWindow('OCR')

                else:
                    print("Already scanned! Move onto next beacon!\n")
                    counter -= 1
                    break

    cv2.destroyAllWindows()
except:
    beep.beep(3)
    print("Camera disconnected, workbook saved!\n")
    workbook.close()
