# QR Code & OCR scanner v1.0
#
# Automatically scans QR codes; press 'Q' to capture OCR text
# Once OCR is done, scanning the same QR code shows a magenta tag in the live preview (indicating that it has already been scanned)
# Press 'P' to correct the beacon name if OCR fails (hotkey can be changed)
# Press 'N' to save the workbook ONLY AT THE END OF THE BATCH (script will have to be restarted after this; hotkey can be changed)
#
# Download an IP webcam app to use your phone camera (IP Webcam: https://play.google.com/store/apps/details?id=com.pas.webcam&hl=en&gl=US)
# Set video resolution to 640x480, and orientation to 'portrait'
#
# Melroy Caeiro, 200921-1137

import cv2
import numpy as np
from pyzbar.pyzbar import decode
import easyocr
import cv2
import xlsxwriter
from time import time, ctime
import keyboard
#import warnings
#warnings.filterwarnings("ignore")

t = time()

workbook = xlsxwriter.Workbook('ScanMelroy-{}.xlsx'.format(t))  # Excel workbook name; {} will autofill with the number of seconds since epoch
worksheet = workbook.add_worksheet()
worksheet.write('A1', 'QR Code')
worksheet.write('B1', 'Beacon Name')

keyboard.add_hotkey("n", lambda: workbook.close())      # Hotkey to save the Excel workbook; press only when you're actually done with all scans (need to restart script once pressed)

url = 'http://192.168.2.86:8080/video'      # Phone camera IP address
cap = cv2.VideoCapture(url)     # '0' for PC webcam, 'url' for IP webcam

currentData = b'bruh'
substring = 'B'
text = '?'

workStatus = True
ocrStatus = False
excelStatus = False
editStatus = False
counter = 1

keyboard.add_hotkey("p", lambda : worksheet.write('B{}'.format(counter-1), input("Correct Beacon Name: ")))     # Hotkey to edit the current beacon name entry, if OCR fails

try:
    while (workStatus == True):
        counter += 1
        qrStatus = False

        while (qrStatus == False):
            success, img = cap.read()
            for barcode in decode(img):
                myData = barcode.data.decode('UTF-8')

                if (myData != currentData):  # if myData (QR Code) is new; not scanned before
                    primaryColor = (255, 0, 255)
                    print(myData)
                    currentData = myData  # To ensure we remain on the current QR code
                    ocrStatus = True
                else:  # if myData (QR Code) has already been scanned -> move to OCR
                    ocrStatus = False  # OCR can start now

                pts = np.array([barcode.polygon], np.int32)
                pts = pts.reshape((-1, 1, 2))
                cv2.polylines(img, [pts], True, primaryColor, 5)
                pts2 = barcode.rect
                cv2.putText(img, myData, (pts2[0], pts2[1]), cv2.FONT_HERSHEY_SIMPLEX, 0.9, primaryColor, 2)

            if (ocrStatus == True):
                qrStatus = True  # QR done
                worksheet.write('A{}'.format(counter), myData)
                break
            else:
                cv2.imshow('QR Code', img)
                cv2.waitKey(1)
                qrStatus = False

        while (ocrStatus == True):
            success, img = cap.read()
            cv2.imshow('QR Code', img)
            key = cv2.waitKey(33)

            if key == 'q':
                break
            elif key != -1:
                # img2 = cv2.imshow('OCR', img)
                # cv2.waitKey(1)
                reader = easyocr.Reader(['en'], gpu=True)
                result = reader.readtext(img)
                for detection in result:
                    text = detection[1]
                print(text, "\n")
                worksheet.write('B{}'.format(counter), text)
                ocrStatus = False
                # cv2.destroyWindow('OCR')

        workStatus = True
        cv2.destroyAllWindows()
except:
    workbook.close()
