# -*- coding: utf-8 -*-
"""
NA CZYSTO
"""

import imapclient
import imaplib
import pyzmail
import openpyxl
import datetime
import os
from openpyxl.styles import Alignment

flaga = False
flaga1 = False
flaga2 = False
imaplib._MAXLINE = 1000000
while not flaga:
    try:
        dostawca = input(f"Wybierz dostawce uslug z listy lub jesli twojego dostawcy nie ma na liscie, wpisz dane dostawcy uslug:\n-Gmail: imap.gmail.com \n-Verizon: incoming.verizon.net \n-Interia: poczta.interia.pl \n-Onet: imap.poczta.onet.pl \n-Wirtualna Polska: imap.wp.pl port (993) \n-Outlook.com/Hotmail.com: imap-mail.outlook.com \n-Yahoo Mail: imap.mail.yahoo.com \n-Comcast: imap.comcast.net \n\nWpisz ponizej:\n")
        imapObj = imapclient.IMAPClient(dostawca, ssl = True)
        flaga = True
    except:
        print ("Niepoprawny dostawca. Sprobuj jeszcze raz.")
        continue
    
print(30*" ")
print(30*"*")
print(30*" ")
    
while not flaga1:
    try:
        email = input (f"Podaj adres email: ")
        haslo = input (f"Podaj haslo: ")
        imapObj.login(email, haslo)
        flaga1 = True
    except:
        print ("Nieprawidłowe dane. Sprobuj jeszcze raz.")
        continue

print(30*" ")
print(30*"*")
print(30*" ")


lista_tupli_folderow = imapObj.list_folders()
lista_folderow = [i[2] for i in lista_tupli_folderow if i[2] != "[Gmail]"]
print("Wybierz folder sposrod: ")
for i in lista_folderow:
    print (i)

while not flaga2:
    try:
        folder = input (f"Wpisz nazwe folderu z wyswietlonej listy: ")
        imapObj.select_folder(folder, readonly=True)
        flaga2 = True
    except:
        print ("Nieprawidlowy folder. Przepisz dokladnie (z zachowaniem wielkich liter itp.) jeden z folderow wymienionych na liscie.")
        continue

print(30*" ")
print(30*"*")
print(30*" ")

UIDs = imapObj.search(["ALL"])
UIDsLIST = list (imapObj.search(["ALL"]))
rawMessages = imapObj.fetch(UIDs, ["BODY[]"])
wb = openpyxl.Workbook()
wb.sheetnames
sheet = wb.active
sheet.title = "Emaile"

litery_do_iteracji = ("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M")
for i in range (2, len (UIDs)+2):
    for j in litery_do_iteracji:
        sheet[j+str(i)].alignment = Alignment(vertical='center')
for i in litery_do_iteracji:
    sheet[i+"1"].alignment = Alignment(horizontal='center')
for i in range (1, len (UIDs)+2):
    sheet["E"+str(i)].alignment = Alignment(wrapText=True)

c = sheet["A2"]
sheet.freeze_panes = c
sheet.sheet_view.zoomScale = 60
sheet.column_dimensions['A'].width = 40
sheet.column_dimensions['B'].width = 40
sheet.column_dimensions['C'].width = 50
sheet.column_dimensions['D'].width = 60
sheet.column_dimensions['E'].width = 300

data_zapisu = str(datetime.datetime.today())
data2_zapisu = data_zapisu.replace(":", ".")

ilosc_maili = []
for i in range (1, len (UIDsLIST)+1):
    ilosc_maili.append(i)

def generator_1(liczby):
    lista = list (liczby)
    for i in lista:
        yield i
gen1 = generator_1(ilosc_maili)

def generator_2(liczby):
    lista = list (liczby)
    for i in lista:
        yield i
gen2 = generator_2(ilosc_maili)

def generator_3(liczby):
    lista = list (liczby)
    for i in lista:
        yield i
gen3 = generator_3(ilosc_maili)

def generator_4(liczby):
    lista = list (liczby)
    for i in lista:
        yield i
gen4 = generator_4(ilosc_maili)

def generator_5(liczby):
    lista = list (liczby)
    for i in lista:
        yield i
gen5 = generator_5(ilosc_maili)

for ii in UIDsLIST:
    message = pyzmail.PyzMessage.factory(rawMessages[ii] [b'BODY[]'])      
    temat = message.get_subject().rstrip()
    tresc_tekst = message.text_part.get_payload().decode(message.text_part.charset).rstrip()
    tresc_htlm = message.html_part.get_payload().decode(message.html_part.charset).rstrip()
    nadawca = str(message.get_addresses('from'))
    a = message.get_addresses('from')
    for kk in a:
        nazwa_nadawcy = kk[0]
        email_nadawcy = kk[1]
    data = message.get_decoded_header("date")
    sheet["A1"]="Data"
    sheet["B1"]="Nazwa nadawcy"
    sheet["C1"]="Email Nadawcy"
    sheet["D1"]="Temat emaila"
    sheet["E1"]="Tresc_emaila"
    #for ilosc_wierszy in range (2, len(UIDs)+1):
    sheet["A"+str(next(gen1)+1)]=data
    sheet["B"+str(next(gen2)+1)]=nazwa_nadawcy
    sheet["C"+str(next(gen3)+1)]=email_nadawcy
    sheet["D"+str(next(gen4)+1)]=temat
    sheet["E"+str(next(gen5)+1)]=tresc_tekst

wb.save('EMAILE_Z_DNIA' + data2_zapisu +'.xlsx')

komunikat = ("Sukces. \nPlik zapisany w folderze: \n{} \npod nazwa: \n{}").format(os.getcwd(), 'emaile z dnia ' + data2_zapisu +'.xlsx')
print (komunikat)

# del(a,c, data, data2_zapisu,data_zapisu,dostawca,email,\
#     email_nadawcy,flaga, flaga1, flaga2, folder, haslo,\
#     i, imapObj, j, komunikat, lista_folderow, lista_tupli_folderow, \
#     litery_do_iteracji, message, nadawca, nazwa_nadawcy, rawMessages, \
#     sheet, temat, tresc_htlm, tresc_tekst, UIDs, UIDsLIST, wb)

#%%
