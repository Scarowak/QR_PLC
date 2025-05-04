# -*- coding: utf-8 -*-
"""
Created on Fri Apr 25 20:21:46 2025

@author: Oskar
"""
import snap7
from snap7.util import get_bool
from snap7 import Area
import time 
import pandas as pd
import qrcode
import win32print 
import win32ui 
from PIL import Image, ImageWin, ImageDraw, ImageFont
from enum import Enum


#############  ZMIENNE - UZUPEŁNIJ  ###############


#ADRES IP PLC, STRING
IP = '192.168.191.2'  
#RACK NUMBER
RACK = 0  
#SLOT NUMBER            
SLOT = 0
#DB_BLOCK_NUMBER  
DB_NUMBER_STR=201
#START_BYTE
START_BYTE_STR=66
#SIZE
SIZE_STR=50

DB_NUMBER_ISOUT=210
BYTE_INDEX=136
BIT_INDEX = 2



SAVE_DIR = "C:/Users/Oskar/Desktop/DB_SCRAPPING/"





#plik csv i ostatni kod QR (nie zmieniac)
SAVE_FILE = SAVE_DIR+"DB_BLOCK_STRINGS.csv"
SAVE_IMG = SAVE_DIR+"Last_QR.png"




#Funkcje


def disconnect_plc():
    """
    Rozłącza program z plc
    
    Parametry:
        NONE
        
    Returns:
        NONE
    """
    
    client.disconnect()


def check_str(db_number, start_byte, size):
    '''
    Odczytuje str z podanego data bloku
    
    Parametry:
        db_number - int, numer data bloku
        start_byte - int, bit początkowy
        size - int, wielkosc stringa
        
    Returns:
        data - str, wartosc stringa z podanego bloku DB
    '''
    
    data = client.db_read(db_number, start_byte, size)
    data = str(data)
    data = data[14:-14]
    
    return data
    

def is_out(db_number, byte_index, bit_index):
    
    '''
    Odczytuje dane z podanego data bloku, informuje czy mamy wykonanego ina czy outa
    
    Parametry:
        db_number - int, numer data bloku
        start_byte - int, bit początkowy
        size - int, wielkosc stringa
        
    Returns:
        True/False - bool, True - jeżeli mamy doczynienia z OUT, False jeżeli IN
    '''
    
    

    data = client.read_area(Area.DB, db_number, byte_index, 1)
    
    value = get_bool(data, 0, bit_index)
    
    return value  




def save_str_to_csv(string, out):
    print("ZAPISYWANIE STRINGA DO PLIKU CSV")
    
    file = pd.read_csv(SAVE_FILE)
    
    if out:
        inout = "OUT"
    else:
        inout = "IN"
    
    new_row = pd.DataFrame({"STR1": [string], 
                            "IN / OUT" : [inout]})  
    
    file = pd.concat([file, new_row], ignore_index=True)   
    file.to_csv(SAVE_FILE, index=False)  # index=False, aby uniknąć zapisywania indeksów


    
    ###





def save_qr(string, isout):
    # Generate QR code image
    qr_img = qrcode.make(string).convert("RGB")
    qr_width, qr_height = qr_img.size
    
    # Choose the text based on the flag
    text = "O" if isout else "I"
        
        
    # Load a font (you can specify a .ttf path or use default)
    try:
        font = ImageFont.truetype("arial.ttf", 40)  # Use a common font
    except IOError:
        font = ImageFont.load_default()

    # Estimate text size
    bbox = font.getbbox(text)
    text_width = bbox[2] - bbox[0]
    text_height = bbox[3] - bbox[1]
    
    # Create a new image with space for the QR code and the text
    new_width = qr_width + text_width + 20  # 20 pixels padding
    new_height = max(qr_height, text_height + 20)
    combined_img = Image.new("RGB", (new_width, new_height), "white")

    # Paste QR code on the left
    combined_img.paste(qr_img, (0, 0))

    # Draw the text on the right side
    draw = ImageDraw.Draw(combined_img)
    text_x = qr_width + 10  # 10 pixels padding
    text_y = (new_height - text_height) // 2
    draw.text((text_x, text_y), text, font=font, fill="black")

    # Save the result
    combined_img.save(SAVE_IMG)

    ###
    
def print_qr():
        
    # Open the image
    file_name = 'C:/Users/Oskar/Desktop/DB_SCRAPPING/Last_QR.png'
    
    
    
    
    #PHYSICALWIDTH = 110
    #PHYSICALHEIGHT = 111
    printer_name = win32print.GetDefaultPrinter ()
    hDC = win32ui.CreateDC ()
    hDC.CreatePrinterDC (printer_name)
    #printer_size = hDC.GetDeviceCaps (PHYSICALWIDTH), hDC.GetDeviceCaps (PHYSICALHEIGHT)
    
    bmp = Image.open (file_name)
    if bmp.size[0] < bmp.size[1]:
      bmp = bmp.rotate (90)
    
    hDC.StartDoc (file_name)
    hDC.StartPage ()
    
    dib = ImageWin.Dib (bmp)
    dib.draw (hDC.GetHandleOutput (), (0,0,32*7,57*7))
    
    hDC.EndPage ()
    hDC.EndDoc ()
    hDC.DeleteDC ()





#Połączenie ze sterownikiem
connected = False

while not connected:
    
    try:

        client = snap7.client.Client()


        client.connect(IP, RACK, SLOT)  # IP, Rack, Slot

        connected = True

    except:
        print("Nie udało się połączyć")
        time.sleep(1)




#Inicjacja zmiennej old_str (poprzedni string)
old_str='pusty_string'

#Główna pętla
while True:
    
        try:
            # Odczytuje wartosci
            value =   check_str(DB_NUMBER_STR,START_BYTE_STR,SIZE_STR) 
            out_or_in = is_out(DB_NUMBER_ISOUT,BYTE_INDEX,BIT_INDEX)
            
            if value != old_str:
                save_str_to_csv(value, out_or_in)
                save_qr(value, out_or_in)
                time.sleep(1)
                print_qr()
    
    
    
    
            old_str = value
            
        except:
            print("Blad odczytu danych, sprawdź poprawnosc wpisanego adresu Databloku i Stringa")
        finally:
            print("Sleep 1")
            time.sleep(1)
            print("Sleep 2")
            time.sleep(1)
            

    





    






