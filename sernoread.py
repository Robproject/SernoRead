import tkinter
from tkinter import filedialog
import cv2
import pytesseract
import os
import numpy as np
import openpyxl

# select paths
def get_file():
    global xlsx_path
    xlsx_path = filedialog.askopenfilename()

def get_dir():
    global dir_path
    dir_path = filedialog.askdirectory(title='Select Folder')

def get_outline(image):
    contours,hierarchy = cv2.findContours(image,cv2.RETR_EXTERNAL,cv2.CHAIN_APPROX_SIMPLE)
    # only use main contour
    threshold_area = 3000000
    for cnt in contours:
        area = cv2.contourArea(cnt)
        if area > threshold_area:
            x, y, w, h = cv2.boundingRect(cnt)
            return x,y,w,h

def fix_orientation(img):
    # orient vertically if not
    if img.shape[1] > img.shape[0]:
        img = cv2.rotate(img, cv2.ROTATE_90_CLOCKWISE)
    # crop spot and get avg color
    crop = img[1675:1825, 140:310]
    average = np.average(crop)
    # rotate 180 if flipped
    if average > 100:
        img = cv2.rotate(img, cv2.ROTATE_180)
    return img

def process_photo(photo_path):
    img = cv2.imread(photo_path,0)
    # process img then crop board and orient
    img = cv2.medianBlur(img, 5)
    thresh = cv2.threshold(img, 100, 255, cv2.THRESH_BINARY_INV)[1]
    x,y,w,h = get_outline(thresh)
    crop = img[y:y+h,x:x+w]
    img = fix_orientation(crop)
    return img

def get_sernos(log_path, photos_path):
    for filename in os.listdir(photos_path):
        f = os.path.join(photos_path, filename)
        # check if file
        if os.path.isfile(f):
            if "INPUT" in os.path.basename(f):
                img = process_photo(f)
                # crop serno area, thresh, scale
                img = img[70:140, 875:1020]
                img = cv2.threshold(img, 150, 255, cv2.THRESH_BINARY_INV)[1]
                img = cv2.resize(img, (0, 0), fx=5, fy=5)
                # get text
                text = pytesseract.image_to_string(img, config="-c tessedit_char_whitelist=0123456789 --psm 8") 

def get_xlsxdata(xlsx_path):
    ...
    return xlsx_dict

def compare_data(xlsx_dict, serno_dict):
    ...
    return write_data

def write_xlsx(write_data, xlsx_path):
    ...


if __name__ == '__main__':
    rootwindow = tkinter.Tk()
    rootwindow.title("WLM300 SerNo OCR")
    rootwindow.geometry("350x275")

    xlsx_path = '/'
    dir_path = '/'

    button_getfile = tkinter.Button(rootwindow, text="select log.xlsx", command = get_file)
    button_getdir = tkinter.Button(rootwindow, text ="select photo dir", command = get_dir)
    button_processphotos = tkinter.Button(rootwindow, text="Copy SerNos", command = lambda: get_sernos(xlsx_path, dir_path))

    button_getfile.grid(row=0,column=0,padx=10,pady=10)
    button_getdir.grid(row=0,column=1,padx=10,pady=10)
    button_processphotos.grid(row=1,column=0,padx=10,pady=10)

    rootwindow.mainloop()
    
