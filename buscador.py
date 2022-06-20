import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter.messagebox import showerror
import pandas as pd
import os, subprocess
import webbrowser

class Window:
    def __init__(self, master):
        self.master = master
        master.title("Buscador de documentos")
        master.geometry('320x360')
        master.eval('tk::PlaceWindow . center')

        # Label e input file de directorio
        self.frame = tk.Frame(master)
        self.frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=10)

        self.labelDirectory = tk.Label(
            self.frame, text="Directorio de busqueda")
        self.labelDirectory.place(x=0, y=0, anchor='nw')
        self.labelDirectory.pack(fill=tk.X)

        self.entryDirectory = tk.Entry(self.frame)
        self.entryDirectory.pack(fill=tk.X)

        self.buttonDirectory = tk.Button(
            self.frame, text="Seleccionar", command=self.getDir)
        self.buttonDirectory.pack(fill=tk.X)

        # Label e input file de archivo
        self.frameFile = tk.Frame(master)
        self.frameFile.pack(fill=tk.X, padx=10, pady=10)

        self.labelFile = tk.Label(
            self.frameFile, text="Archivo excel")
        self.labelFile.place(x=0, y=0, anchor='nw')
        self.labelFile.pack(fill=tk.X)

        self.entryFile = tk.Entry(self.frameFile)
        self.entryFile.pack(fill=tk.X)

        self.buttonFile = tk.Button(
            self.frameFile, text="Seleccionar", command=self.getFile)
        self.buttonFile.pack(fill=tk.X)

        # search input
        self.frameSearch = tk.Frame(master)
        self.frameSearch.pack(fill=tk.X, padx=10, pady=10)

        self.labelSearch = tk.Label(
            self.frameSearch, text="Input busqueda")
        self.labelSearch.place(x=0, y=0, anchor='nw')
        self.labelSearch.pack(fill=tk.X)

        self.entrySearch = tk.Entry(self.frameSearch)
        self.entrySearch.pack(fill=tk.X)

        self.buttonSearch = tk.Button(
            self.frameSearch, text="Buscar", command=self.initSearch)
        self.buttonSearch.pack(fill=tk.X)

        # lista de concurrencias
        self.frameList = tk.Frame(master)
        self.frameList.pack(fill=tk.X, padx=10, pady=10)
        self.listbox = tk.Listbox(self.frameList)
        self.listbox.pack(fill=tk.X)
        self.listbox.bind('<Double-1>', self.sendToPrint) 

    def getFile(self):
        filename = filedialog.askopenfilename(
            initialdir="/", title="Select a File", filetypes=(("Text files", "*.xlsx*"), ("all files", "*.*")))
        if(filename != ""):
            self.entryFile.insert(0, filename)

    def getDir(self):
        directory = filedialog.askdirectory()
        if directory != "":
            self.entryDirectory.insert(0, directory)

    def initSearch(self):
        filePath = ""
        self.stateWindow(tk.DISABLED)

        if(len(self.entryDirectory.get()) <= 0):
            showerror("Error", "No ha seleccionado directorio")
            self.stateWindow(tk.NORMAL)
            return

        if(len(self.entryFile.get()) <= 0):
            showerror("Error", "No ha seleccionado archivo")
            self.stateWindow(tk.NORMAL)
            return

        filePath = self.entryFile.get()
        originDirectory = self.entryDirectory.get()
        excelContent = pd.read_excel(filePath)
        self.clearList()

        searchDF = excelContent[excelContent['Nombre'].str.contains(self.entrySearch.get())]
        if(searchDF.empty):
            searchDF = excelContent[excelContent['Id Empleado'].astype('str').str.contains(self.entrySearch.get())]

        if(searchDF.empty):
            self.stateWindow(tk.NORMAL)
            return 

        for i in searchDF.index:
            filePath = ''
            filePath = self.findFile(originDirectory, searchDF['RFC'][i] + ".pdf")
            if(filePath != '') :
                self.listbox.insert(0, filePath)
       
        self.stateWindow(tk.NORMAL)

    def stateWindow(self, state):
        self.entryDirectory['state'] = state
        self.entryFile['state'] = state
        self.buttonDirectory['state'] = state
        self.buttonFile['state'] = state

    def findFile(self, path, name):
        print(name, path)
        for dirpath, dirname, filename in os.walk(path):
            if name in filename:
                return dirpath+"/"+name

    def clearList(self):
        self.listbox.delete(0,tk.END)

    def sendToPrint(self, event):
        cs = self.listbox.get(self.listbox.curselection()) 
        webbrowser.open_new(cs)

root = tk.Tk()
window = Window(root)
root.mainloop()
