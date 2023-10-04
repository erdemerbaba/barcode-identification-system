# -*- coding: utf-8 -*-

import cv2
from pyzbar import pyzbar
import xlwings as xw
from openpyxl import load_workbook

def read_barcodes(frame):
    barcodes = pyzbar.decode(frame)
    for barcode in barcodes:
        x, y , w, h = barcode.rect
        barcode_info = barcode.data.decode('utf-8')
        cv2.rectangle(frame, (x, y),(x+w, y+h), (0, 255, 0), 2)
        font = cv2.FONT_HERSHEY_DUPLEX
        cv2.putText(frame, barcode_info, (x + 6, y - 6), font, 2.0, (255, 255, 255), 1)
        with open("barkod.txt", mode ='w') as file:
            file.write("Tanınan Barkod:" + barcode_info)
            print(int(barcode_info))
            bookName = 'barkod.xlsx'
            sheetName = 'Sheet1'
            wb = xw.Book(bookName)
            myCell = wb.sheets[sheetName].api.UsedRange.Find(int(barcode_info))
            print (myCell.address)
            wb = load_workbook(filename = 'barkod.xlsx')
            sheet_ranges = wb['Sheet1']
            print(sheet_ranges["B"+myCell.address[3]].value)
            print("  ")
    return frame

def main():
    camera = cv2.VideoCapture(0)
    ret, frame = camera.read()
    while ret:
        ret, frame = camera.read()
        frame = read_barcodes(frame)
        cv2.imshow('Barkod/QR kod okuyucu', frame)
        if cv2.waitKey(1) & 0xFF == 27:
            break
    camera.release()
    cv2.destroyAllWindows()

if __name__ == '__main__':
    main()
