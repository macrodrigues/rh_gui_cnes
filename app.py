""" This script launches the GUI """
#pylint: disable-msg=too-many-arguments
import os
import tkinter as tk
from tkinter import messagebox as msg
from tkinter import ttk, Entry, Frame, Label, Button, Tk, LabelFrame
from PIL import Image, ImageTk
from extract_mails import extract, write, get_email_accounts

#ouput path
DESKTOP = os.path.join( # Desktop path for 'this' user
    os.path.join(os.environ['USERPROFILE']),
    'Desktop')

EXCEL_NAME = "Candidatures.xlsx"

class App:
    """ Main class """
    def __init__(self):
        """When the GUI is opened, the instructions below are implemented"""
        self.win = Tk() #launch tkinter GUI
        self.img_tk = None # varibales defined in init for grood practice
        self.label_img = None
        self.account_frame = None
        self.lbl_account = None
        self.account_entry = None
        self.subfolder_frame = None
        self.lbl_subfolder = None
        self.subfolder_entry = None
        self.excel_frame = None
        self.lbl_excel = None
        self.excel_entry = None
        self.wait_label = None
        self.main_button = None
        self.win.configure(background='#004C99') # background color
        self.win.title('Gestionnaire des Bourses') # title
        self.win.iconbitmap("C:\\Users\\Marco Rodrigues\\Desktop\\script\\outils\\cnes.ico") # call icon
        self.win.minsize(480,625) # min size of window
        self.win.maxsize(480,625) # max size of window
        self.set_description() # set description frame
        self.set_main_frame()
        self.set_main_page()

    def set_description(self):
        """Displays a short descriptions at the beggining"""
        self.intro_frame = tk.Frame(self.win, bg='#004C99')
        self.intro_frame.pack()
        text = \
        """Ce script télécharge les formulaires DERC, DEX et DARC.
Les informations sont extraites et placées dans un Excel."""
        self.label_intro = tk.Label(
            self.intro_frame, 
            text=text, 
            bg='#004C99',
            font='Arial 10 bold',
            fg = 'white')
        self.label_intro.pack(side='top', pady=5)

    def set_main_frame(self):
        """Creates main frame, where the widgets will be displayed"""
        self.main_frame = LabelFrame(
            self.win,
            relief=tk.SUNKEN,
            bg='white')
        self.main_frame.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

    def set_main_page(self):
        """Widgets of the first page displayed, by running this function"""
        self.set_background_image()
        self.set_account()
        self.set_subfolder()
        self.set_excel_entry()
        self.set_main_button()

    def set_background_image(self):
        """Prettify GUI with a background image"""
        img_init = Image.open("C:\\Users\\Marco Rodrigues\\Desktop\\script\\outils\\academie_guyane.PNG")
        img_init = img_init.resize((300, 180)) # image size
        self.img_tk = ImageTk.PhotoImage(img_init)
        self.label_img = tk.Label(
            self.main_frame,
            image = self.img_tk,
            borderwidth=0)
        self.label_img.pack()

    def set_account(self):
        """This function allows to choose the outlook account,
        from where the data will be extracted"""
        self.account_frame = Frame(self.main_frame, bg='white')
        self.account_frame.pack(pady = 5, padx = 10)
        self.lbl_account = Label(
            self.account_frame,
            bg='white',
            text='Choisissez votre compte Outlook:',
            pady = 5,
            font='Arial 10 bold')
        self.account_entry = ttk.Combobox(
            self.account_frame,
            values = get_email_accounts(),
            width=40,
            state='readonly')
        try:
            self.account_entry.current(1)
        except Exception as eee:
            print(eee)
            self.account_entry.current(0)
        self.lbl_account.pack()
        self.account_entry.pack()

    def set_subfolder(self):
        """Choose the inbox, by default is 'Boîte de reception'"""
        self.subfolder_frame = Frame(self.main_frame, bg='white')
        self.subfolder_frame.pack(pady = 5)
        self.lbl_subfolder = Label(
            self.subfolder_frame,
            text='Spécifiez la Boîte Outlook:',
            bg='white',
            pady = 5,
            font='Arial 10 bold')
        self.subfolder_entry = Entry(self.subfolder_frame,
                              bd=5,
                              width=30)
        self.subfolder_entry.insert(-1, 'Boîte de reception')
        self.lbl_subfolder.pack()
        self.subfolder_entry.pack()

    def set_excel_entry(self):
        """Defines where to save the excel file, it saves in desktop by default"""
        self.excel_frame = Frame(self.main_frame, bg='white')
        self.excel_frame.pack(pady = 5)
        self.lbl_excel = Label(
            self.excel_frame,
            text='Spécifiez la destination du fichier Excel:',
            pady = 5,
            bg='white',
            font='Arial 10 bold')
        self.excel_entry = Entry(self.excel_frame,
                        bd=5,
                        width=30)
        self.excel_entry.insert(0, os.path.join(DESKTOP))
        self.lbl_excel.pack()
        self.excel_entry.pack()

    def loading_component(self):
        """Make a loading window"""
        self.wait_label = tk.Label(
            text= 'En cours...',
            bg= '#004C99',
            font='Arial 15 bold',
            fg = 'white',
            borderwidth=0)
        self.wait_label.pack(side='bottom', pady=20)

    def set_main_button(self):
        """Button widget"""
        self.main_button = Button(self.main_frame,
                                   text='Sélectionner',
                                   font='Helvetica 11 bold',
                                   command=self.action_main_button,
                                   bg='#004C99',
                                   activebackground='#0000CD',
                                   fg='#FFFFFF')
        self.main_button.pack(pady = 10)
  
    def action_main_button(self):
        """Callback function for the button"""
        excel_path = os.path.join(self.excel_entry.get(), EXCEL_NAME)
        bool_temp = os.path.exists(self.excel_entry.get())
        if not bool_temp: # if False the path does not exist
            msg.showerror('Erreur', "Cette destination n'existe pas")
        self.loading_component()
        try:
            if os.path.exists(excel_path):
                res = msg.askyesno('Info','Ce fichier existe déjà, êtes-vous sûr de vouloir le remplacer?')
                if res:
                    account = self.account_entry.get()
                    subfolder = self.subfolder_entry.get()
                    headers, values = extract(account, subfolder)
                    write(headers, values, excel_path)
                    self.wait_label.destroy()
                    msg.showinfo('Info', 'Le fichier à été bien enregistré!')
                self.win.destroy() # quit app after message
    
            else:
                account = self.account_entry.get()
                subfolder = self.subfolder_entry.get()
                headers, values = extract(account, subfolder)
                write(headers, values, excel_path)
                self.wait_label.destroy()
                msg.showinfo('Info', 'Le fichier à été bien enregistré!')
                self.win.destroy() # quit app after message
        except Exception as excep:
            if type(list(excep.args)[0]) == PermissionError:
                msg.showerror('Erreur', "Veulliez fermer l'excel.")
            else:
                self.wait_label.destroy()
                msg.showerror(
                    'Erreur',
                    "Une erreur s'est produite")
                self.win.destroy()

app = App()
app.win.mainloop()