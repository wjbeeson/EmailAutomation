import tkinter
import tkinter.filedialog
from datetime import datetime
from pathlib import Path
from tkinter import *
from tkinter import filedialog
from tkinter import simpledialog
from tkinter import messagebox

import main_form
from validate_sheet import validate_csv_file
from account_management import *
from main_form import *

class AccountMenu(Frame):
    def __init__(self, root):
        for ele in root.winfo_children():
            ele.destroy()
        self.base = root
        super().__init__()
        self.color = "Green"

        self.base.geometry('200x250')
        self.base.title("Login")
        current_row = 0

        self.back_button = Button(self.base, text='Back to Menu', bg=self.color, fg='white', command=self.back)
        self.back_button.grid(
            row=current_row,
            column=0,
            columnspan=4,
            sticky=W + E)
        current_row += 1

        # Username
        Label(self.base, text="Email: ", font=("bold", 10)).grid(row=current_row, column=0, columnspan=1,
                                                                    sticky=W)
        self.username = tkinter.StringVar()
        Entry(self.base, textvariable=self.username, fg='black'). grid(row=current_row, column=1, columnspan=1, sticky=W + E)
        current_row += 1
        # Password
        Label(self.base, text="Password: ", font=("bold", 10)).grid(row=current_row, column=0, columnspan=1,
                                                                    sticky=W)
        self.password = tkinter.StringVar()
        Entry(self.base, textvariable=self.password, fg='black'). grid(row=current_row, column=1, columnspan=1, sticky=W + E)
        current_row += 1

        self.try_login_button = Button(self.base, text='Try Login', bg=self.color, fg='white', command=self.try_add_account)
        self.try_login_button.grid(
            row=current_row,
            column=0,
            columnspan=4,
            sticky=W + E)
        current_row += 1

        self.error_message_text = tkinter.StringVar()
        self.error_message = Label(self.base, textvariable=self.error_message_text, font=("bold", 10), foreground="red", justify='left')
        current_row += 1

        self.error_message.grid(row=current_row, column=0, columnspan=4, sticky=E)
        self.set_form_values("","","")

        self.base.mainloop()

    def set_form_values(self, error_text=None, username=None, password=None):
        if error_text is not None:
            self.error_message_text.set(error_text)
        if username is not None:
            self.username.set(username)
        if password is not None:
            self.password.set(password)

    def back(self):
        main_form.MainMenu(self.base)

    def try_add_account(self):
        if not test_valid_credentials(self.username.get(), self.password.get()):
            print("incorrect pass")
            self.set_form_values(
                error_text="Login failed. Please do the \n"
                           "following:\n\n"
                           "1. Check to make sure 'less\n"
                           "secure app access' is enabled\n"
                           "for your account.\n\n"
                           "2. Double check your login\n"
                           "Information.\n",
                password=""
            )
            return
        print("correct pass")
        add_account(self.username.get(), self.password.get())


