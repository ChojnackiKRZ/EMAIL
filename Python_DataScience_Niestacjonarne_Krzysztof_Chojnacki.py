# -*- coding: utf-8 -*-
"""

@author: krzysztof
"""
"""
Wszystkie moduły firm trzecich używane w projekcie można zainstalować 
wydając wymienione poniżej polecenia w konsoli. Jeśli używasz systemu 
operacyjnego macOS lub Linux, wówczas pip zastąp przez pip3.
pip install openpyxl
pip install imapclient
pip install pyzmail36

"""

#pobor odpowiednich modulow
import imapclient
import imaplib
import pyzmail
import openpyxl
import datetime
from openpyxl.styles import Alignment
import tkinter
from tkinter import ttk
import os

root=tkinter.Tk()

#ta funkcja sluzy do tego, aby moc wywolywac wiele funkcji w przycisku
#start

def lacz_funkcje(*funcs):
    def zlaczone_funkcje(*args, **kwargs):
        for f in funcs:
            f(*args, **kwargs)
    return zlaczone_funkcje

#tutaj mam zdefiniowane funkcje dotyczace wartosci oraz akcji w okienach
#dla poszczegolnych czynnosci: pobieranie emaila, hasla, klienta poczty
#oraz logowanie przy ich pomocy
#zdefiniowana rowniez funkcja do wyswietlania wartosci z obiektu
#imapObj.list_folders()
#nastepnie funkcja zapisz_dane wywoluje czesc kodu skupiona na obrobce
#danych do formatu excela i zapisie ich

#w tej funkcji definiuje pobieranie wartosci z okna okno_email
def email():
    global email
    email = startEntry1.get()
#w tej funkcji definiuje pobieranie wartosci z okna okno_haslo
def haslo():
    global haslo
    haslo = startEntry2.get()
#w tej funkcji definiuje pobieranie wartosci z okna lista_rozwijana_dostawcow
def klient_poczty():
    global klient_poczty
    klient_poczty = lista_rozwijana_dostawcow.get()
#w tej fukncji definiuje wykorzystanie emaila, hasla i klienta poczty
#do logowania, stworzenia obiektow poczty, usuniecia pierwszego guzika
#"logowanie" i utworzenie nowego "pobierz dane"
def logowanie():
    labelTop = tkinter.Label(root,text = "Wybierz skrzynke")
    labelTop.pack()
    comboExample.pack()
    imaplib._MAXLINE = 1000000
    global imapObj, lista_folderow, plotButton2
    imapObj = imapclient.IMAPClient(klient_poczty, ssl = True)
    imapObj.login(email, haslo)
    lista_tupli_folderow = imapObj.list_folders()
    lista_folderow = [i[2] for i in lista_tupli_folderow if i[2] != "[Gmail]"]
    plotButton.destroy()
    plotButton2 = tkinter.Button(root, text = "pobierz dane", command=lacz_funkcje(skrzynka_dane, zapisz_dane))
    plotButton2.pack()
#w tej funkcji definiuje stworzenie listy folderow w skrzynce na podstawie
#odczytu ze skrzynki pocztowej. Foldery nie sa wiec zdefiniowane na sztywno,
#co zaklocaloby dzialanie programu
def nowa_lista():
    comboExample["values"] = lista_folderow
#w tej funkcji definiuje pobieranie wartosci w postaci wybranej lub wpisanego
#folderu dostpenego na poczcie uzytkownika
def skrzynka_dane():
    global skrzynka_dane
    skrzynka_dane = comboExample.get()
    comboExample.pack()
#w tej funkcji definiuje dzialania zwiazane z odczytem danych z obiektow
#poczty, definiuje sposob ich przetworzenia oraz sposob i format zapisu
def zapisz_dane():
    imapObj.select_folder(skrzynka_dane, readonly=True)
    UIDs = imapObj.search(["ALL"])
    UIDsLIST = list (imapObj.search(["ALL"]))
    rawMessages = imapObj.fetch(UIDs, ["BODY[]"])
    wb = openpyxl.Workbook()
    wb.sheetnames
    sheet = wb.active
    sheet.title = "Emaile"
    litery_do_iteracji = ("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", 
                          "K", "L", "M")
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
#obiekt generatordo niezaleznej iteracji po liczniku maili   
    def generator_1(liczby):
        lista = list (liczby)
        global numery_maili
        for numery_maili in lista:
            yield numery_maili
    gen1 = generator_1(ilosc_maili)
#w tej petli for nalezacej do funkcji zapisz_dane() obsluguje wyjatek zwiazany
#z formatem obrazkowym/html wiadomosci. Jesli wiadomosc jest obrazkowa/nie ma
#tekstu, ma iterowac i zapisac po jej formacie HTML   
    for indeksy_maili in UIDsLIST:
        message = pyzmail.PyzMessage.factory(rawMessages[indeksy_maili] [b'BODY[]'])      
        temat = message.get_subject().rstrip()
        try:
            tresc_tekst = message.text_part.get_payload().decode(message.text_part.charset).rstrip()
            nadawca = message.get_addresses('from')
            for nazwa_email in nadawca:
                nazwa_nadawcy = nazwa_email[0]
                email_nadawcy = nazwa_email[1]
            data = message.get_decoded_header("date")
            sheet["A1"]="Data"
            sheet["B1"]="Nazwa nadawcy"
            sheet["C1"]="Email Nadawcy"
            sheet["D1"]="Temat emaila"
            sheet["E1"]="Tresc_emaila"
            sheet["A"+str(next(gen1)+1)]=data
            sheet["B"+str(numery_maili+1)]=nazwa_nadawcy
            sheet["C"+str(numery_maili+1)]=email_nadawcy
            sheet["D"+str(numery_maili+1)]=temat
            sheet["E"+str(numery_maili+1)]=tresc_tekst
        except AttributeError:
            tresc_html = message.html_part.get_payload().decode(message.html_part.charset).rstrip()
            nadawca = message.get_addresses('from')
            for nazwa_email in nadawca:
                nazwa_nadawcy = nazwa_email[0]
                email_nadawcy = nazwa_email[1]
            data = message.get_decoded_header("date")
            sheet["A1"]="Data"
            sheet["B1"]="Nazwa nadawcy"
            sheet["C1"]="Email Nadawcy"
            sheet["D1"]="Temat emaila"
            sheet["E1"]="Tresc_emaila"
            sheet["A"+str(next(gen1)+1)]=data
            sheet["B"+str(numery_maili+1)]=nazwa_nadawcy
            sheet["C"+str(numery_maili+1)]=email_nadawcy
            sheet["D"+str(numery_maili+1)]=temat
            sheet["E"+str(numery_maili+1)]=tresc_html
    
    wb.save('EMAILE_Z_DNIA_' + data2_zapisu +'.xlsx')
    
    komunikat = ("Sukces. \nPlik zapisany w folderze: \n{} \npod nazwa: \n{}").format(os.getcwd(), 'emaile z dnia ' + data2_zapisu +'.xlsx')
    plotButton2.destroy()
    komunikat_okienko = tkinter.Label(root,text=komunikat)
    komunikat_okienko.pack()
    
#tutaj mam zdefiniowane okienka z wartosciami dot. maila, hasla, klientow
#poczty oraz dostepnych skrzynek
okno_email = tkinter.Label(root,text="Wpisz email: ")
okno_email.pack()
startEntry1 = tkinter.Entry(root)
startEntry1.pack()

okno_haslo = tkinter.Label(root,text="Wpisz haslo: ")
okno_haslo.pack()
startEntry2 = tkinter.Entry(root, show = "*")
startEntry2.pack()

okno_klient_poczty = tkinter.Label(root,text="Wybierz lub wpisz klienta poczty: ").pack()
lista_dostawcow = ["imap.gmail.com", "poczta.interia.pl","imap.poczta.onet.pl",
                   "imap.wp.pl","imap-mail.outlook.com", "imap.mail.yahoo.com", 
                   "imap.comcast.net"]
lista_rozwijana_dostawcow = ttk.Combobox(root, values = lista_dostawcow)
lista_rozwijana_dostawcow.pack()

plotButton = tkinter.Button(root,text="LOGOWANIE", command=lacz_funkcje(email, haslo, klient_poczty, logowanie, nowa_lista))
plotButton.pack()

#tutaj zdefioniwana lista rozwijana skrzynek na mailu z postcommandem
#aby były wyswietlane dynamicznie na podstawie wartosci ze skrzynki
lista_rozwijana = []
comboExample = ttk.Combobox(root, values = lista_rozwijana, postcommand = nowa_lista)

root.mainloop()
