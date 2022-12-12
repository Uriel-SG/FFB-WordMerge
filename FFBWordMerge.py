import tkinter as tk
from tkinter import BOTTOM, TOP, Menu
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo
from tkinter.simpledialog import askstring
from unicodedata import name
from docxcompose.composer import Composer
from docx import Document
import os
import webbrowser
import pyautogui as pa
import time

#La nostra finestra
window = tk.Tk()
window.title("FFB-Word-Merging")
window.geometry("1599x899")
window.iconbitmap("ffb.ico")
window.attributes("-fullscreen", True)

#Menubar
menubar = tk.Menu(window)
file_menu = tk.Menu(menubar)
file_menu.add_command(label='Exit', command=window.destroy)
menubar.add_cascade(label="File", menu=file_menu)
window.config(menu=menubar)

#Messaggio iniziale
showinfo(title="Benvenuto/a!", message="Caro/a fratello/sorella FFB,\nSei pronto/a ad unire i tuoi file Word?\nIniziamo!")

#Costruzione finestra
title = tk.Label(text="FFB\nUniWord",
fg= "white",
bg= "light blue",
height=3, 
width=1500,
font=("Times", 46, "bold"))
title.pack(fill=tk.X)

sottotitolo1 = tk.Label(
text="il tuo fraterno 'unisci-word' sempre fedele!\n\n",
background="light blue",
foreground="white",
font=("Times", 18, "italic"),
)
sottotitolo1.pack(fill=tk.X)

sottotitolo2 = tk.Label(
text="powered by fra Sal",
background="light blue",
foreground="white",
font=("Times", 15, "bold"))
sottotitolo2.pack(fill=tk.X)

spazio = tk.Label()
spazio.pack(fill=tk.X)

#Logo Word
logo_word = tk.PhotoImage(file="r_microsoft-word-logo.png")
word = tk.Label(image=logo_word)
word.pack()

spazio2 = tk.Label()
spazio2.pack(fill=tk.X)

#File da unire
name = ""

def namefinal():
    number = 1
    global name
    name = askstring('Nome file unito', 'Digita il nome che desideri dare al documento finale: ') + ".docx"
    namepure = name[:-5]
    doclist = os.listdir('C:/WordMergeOutput')
    while name in doclist:
        name = namepure + str(number) + ".docx"
        number += 1
    return name


def select_files():
    filetypes = (
        ('Word files', '*.docx'),
        ('All files', '*.*')
    )
    filenames = fd.askopenfilenames(
        title='Open files',
        initialdir='C:/',
        filetypes=filetypes)
    paths = ['empty.docx']
    for y in filenames:
        paths.append(y)
    dirlist = os.listdir("C:/")
    if 'WordMergeOutput' not in dirlist:
        os.mkdir("C:/WordMergeOutput")
    filename = namefinal()

    master = Document(paths[0])
    composer = Composer(master)
    for x in paths:
        doc = Document(x)
        composer.append(doc)
        composer.save(f"C:/WordMergeOutput/{filename}")
    showinfo(
        title='Selected Files',
        message= f"Fatto! Troverai il documento {filename} in C:/WordMergeOutput/"
    )

    webbrowser.open("C:/WordMergeOutput/") 



# open button
open_button = ttk.Button(
    window,
    text='Open Files',
    command=select_files
)
open_button.pack()

spazio2 = tk.Label(height=7)
spazio2.pack(fill=tk.X)


window.mainloop()




