# -*- coding: utf-8 -*-
"""
Created on Fri Apr 25 20:21:46 2025

@author: Oskar
"""
import snap7
from snap7.util import get_int
import time 
import pandas as pd
import qrcode
import win32print 
import win32ui 
from PIL import Image, ImageWin 


#############  ZMIENNE - UZUPEŁNIJ  ###############


#ADRES IP PLC, STRING
IP = '192.168.171.1'  
#RACK NUMBER
RACK = 0  
#SLOT NUMBER            
SLOT = 0
#DB_BLOCK_NUMBER  
DB_NUMBER=2137
#START_BYTE
START_BYTE=69
#SIZE
SIZE=69
SAVE_DIR = "C:/Users/Oskar/Desktop/DB_SCRAPPING/"




#plik csv i ostatni kod QR (nie zmieniac)
SAVE_FILE = SAVE_DIR+"DB_BLOCK_STRINGS.csv"
SAVE_IMG = SAVE_DIR+"Last_QR.png"

#Inicjacja zmiennej
old_str='pusty_string'


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
    '''
    
    data = client.db_read(db_number, start_byte, size)
    data = str(data)
    
    



def save_str_to_csv(string):
    print("ZAPISYWANIE STRINGA DO PLIKU CSV")
    
    file = pd.read_csv(SAVE_FILE)
    
    new_row = pd.DataFrame({"STR1": [string]})  # Używaj jednej zmiennej new_row
    
    file = pd.concat([file, new_row], ignore_index=True)
    
    file.to_csv(SAVE_FILE, index=False)  # index=False, aby uniknąć zapisywania indeksów


    
    ###




def save_qr(string):
    img = qrcode.make(string)
    img.save(SAVE_IMG)

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



#Główna pętla
while True:
    
        try:
            # Odczytuje wartosci
            value =   check_str(DB_NUMBER,START_BYTE,SIZE)
            
            if value != old_str:
                save_str_to_csv(value)
                save_qr(value)
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
            

    





    






