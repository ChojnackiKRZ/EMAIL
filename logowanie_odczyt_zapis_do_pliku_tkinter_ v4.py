# -*- coding: utf-8 -*-
"""
Created on Mon Jan  4 20:15:16 2021

@author: krzys_npdvhw6
"""

import imapclient
import imaplib
import pyzmail
import openpyxl
import datetime
from openpyxl.styles import Alignment
import tkinter
# from tkinter import ttk
import os

root=tkinter.Tk()
#ta funkcja sluzy do tego, aby laczyc funkcje w przycisku start

def lacz_funkcje(*funcs):
    def zlaczone_funkcje(*args, **kwargs):
        for f in funcs:
            f(*args, **kwargs)
    return zlaczone_funkcje

#tutaj mam zdefiniowane funkcje dotyczace wartosci w okienach
def email():
    global email
    email = startEntry1.get()
def haslo():
    global haslo
    haslo = startEntry2.get()
def klient_poczty():
    global klient_poczty
    klient_poczty = startEntry3.get()
def logowanie():
    labelTop = tkinter.Label(root,text = "Wybierz skrzynke")
    labelTop.pack()
    comboExample.pack()
    imaplib._MAXLINE = 1000000
    global imapObj
    imapObj = imapclient.IMAPClient(klient_poczty, ssl = True)
    imapObj.login(email, haslo)
    lista_tupli_folderow = imapObj.list_folders()
    global lista_folderow
    lista_folderow = [i[2] for i in lista_tupli_folderow if i[2] != "[Gmail]"]
    plotButton.destroy()
    global plotButton2
    plotButton2 = tkinter.Button(root, text = "pobierz dane", command=lacz_funkcje(skrzynka_dane, zapisz_dane))
    plotButton2.pack()
def nowa_lista():
    comboExample["values"] = lista_folderow
def skrzynka_dane():
    global skrzynka_dane
    skrzynka_dane = comboExample.get()
    comboExample.pack()

def zapisz_dane():
    imapObj.select_folder(skrzynka_dane, readonly=True)
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
        sheet["A"+str(next(gen1)+1)]=data
        sheet["B"+str(next(gen2)+1)]=nazwa_nadawcy
        sheet["C"+str(next(gen3)+1)]=email_nadawcy
        sheet["D"+str(next(gen4)+1)]=temat
        sheet["E"+str(next(gen5)+1)]=tresc_tekst
    
    wb.save('EMAILE_Z_DNIA_' + data2_zapisu +'.xlsx')
    
    komunikat = ("Sukces. \nPlik zapisany w folderze: \n{} \npod nazwa: \n{}").format(os.getcwd(), 'emaile z dnia ' + data2_zapisu +'.xlsx')
    print (komunikat)
    
#tutaj mam zdefiniowane okienka z wartosciami
okno_email = tkinter.Label(root,text="Wpisz email: ")
okno_email.pack()
startEntry1 = tkinter.Entry(root)
startEntry1.pack()

okno_haslo = tkinter.Label(root,text="Wpisz haslo: ")
okno_haslo.pack()
startEntry2 = tkinter.Entry(root, show = "*")
startEntry2.pack()

okno_klient_poczty = tkinter.Label(root,text="Wpisz klienta poczty: ").pack()
startEntry3 = tkinter.Entry(root)
startEntry3.pack()

plotButton = tkinter.Button(root,text="LOGOWANIE", command=lacz_funkcje(email, haslo, klient_poczty, logowanie, nowa_lista))
plotButton.pack()
"""
po tym dobrze by bylo kliknac start, nastpenie zeby sie wyswietlila lista
folderow
"""
lista_rozwijana = []
comboExample = ttk.Combobox(root, values = lista_rozwijana, postcommand = nowa_lista)



root.mainloop()