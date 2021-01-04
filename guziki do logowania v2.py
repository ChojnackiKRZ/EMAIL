import imapclient
import imaplib
import pprint
import pyzmail
import openpyxl
import datetime
from openpyxl.styles import Alignment
import tkinter
from tkinter import ttk
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
    skrzynka = comboExample.get()
    labelTop = tkinter.Label(root,text = "Wybierz skrzynke")
    labelTop.pack()
    comboExample.pack()
    imaplib._MAXLINE = 1000000
    imapObj = imapclient.IMAPClient(klient_poczty, ssl = True)
    imapObj.login(email, haslo)
    lista_tupli_folderow = imapObj.list_folders()
    global lista_folderow
    lista_folderow = [i[2] for i in lista_tupli_folderow if i[2] != "[Gmail]"]
def nowa_lista():
    comboExample["values"] = lista_folderow
    
#tutaj mam zdefiniowane okienka z wartosciami
okno_email = tkinter.Label(root,text="Wpisz email: ").pack()
startEntry1 = tkinter.Entry(root)
startEntry1.pack()

okno_haslo = tkinter.Label(root,text="Wpisz haslo: ").pack()
startEntry2 = tkinter.Entry(root)
startEntry2.pack()

okno_klient_poczty = tkinter.Label(root,text="Wpisz klienta poczty: ").pack()
startEntry3 = tkinter.Entry(root)
startEntry3.pack()

plotButton = tkinter.Button(root,text="LOGOWANIE", command=lacz_funkcje(email, haslo, klient_poczty, logowanie, nowa_lista)).pack()
"""
po tym dobrze by bylo kliknac start, nastpenie zeby sie wyswietlila lista
folderow
"""
lista_rozwijana = []
comboExample = ttk.Combobox(root, values = lista_rozwijana, postcommand = nowa_lista)



root.mainloop()