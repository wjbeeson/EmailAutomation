import tkinter.filedialog
from account_form import *
from account_management import *
from send_emails import *
from tkinter import Label
from tkinter import *


class MainMenu:
    def __init__(self, root):
        for ele in root.winfo_children():
            ele.destroy()
        self.base = root
        self.selected_sheet = ""
        self.selected_message = ""
        self.selected_account = ""
        self.color = "Green"


        self.base.geometry('450x400')
        self.base.title("Email Bot")
        current_row = 0

        Label(self.base, text="Account: ", font=("bold", 10)).grid(row=current_row, column=0, columnspan=1,
                                                                            sticky=W)
        self.account_menu_selection = tkinter.StringVar()
        self.account_menu_selection.trace("w", self.option_menu_click)
        #  Dropdown menu to change account
        self.account_menu = OptionMenu(self.base, self.account_menu_selection, *[""])
        self.account_menu.grid(row=current_row, column=1, sticky=EW, columnspan=1)

        current_row += 1

        # Buttons to select and validate a CSV sheet
        Label(self.base, text="Contact Sheet: ", font=("bold", 10)).grid(row=current_row, column=0, columnspan=1,
                                                                            sticky=W)
        self.select_sheet_button_text = tkinter.StringVar()
        self.select_sheet_button = Button(self.base, textvariable=self.select_sheet_button_text, fg='black',
                                          command=self.select_sheet)
        self.select_sheet_button.grid(row=current_row, column=1, columnspan=1, sticky=W + E)
        Button(self.base, text='Validate Sheet', bg=self.color, fg='white', command=self.validate_sheet).grid(
            row=current_row,
            column=2,
            columnspan=1,
            sticky=W + E)
        current_row += 1

        # Select message button
        Label(self.base, text="Message Template: ", font=("bold", 10)).grid(row=current_row, column=0, columnspan=1,
                                                                            sticky=W)
        self.select_message_button_text = tkinter.StringVar()
        self.select_message_button = Button(self.base, textvariable=self.select_message_button_text, fg='black',
                                            command=self.select_message)
        self.select_message_button.grid(
            row=current_row,
            column=1,
            columnspan=1,
            sticky=W + E)
        current_row += 1

        self.send_emails_button = Button(self.base, text='Send Emails', bg=self.color, fg='white',
                                         command=self.send_emails)
        self.send_emails_button.grid(
            row=current_row,
            column=0,
            columnspan=4,
            sticky=W + E)
        current_row += 1

        self.error_message_text = tkinter.StringVar()
        self.error_message = Label(self.base, textvariable=self.error_message_text, font=("bold", 10), foreground="red")
        current_row += 1

        self.error_message.grid(row=current_row, column=0, columnspan=4, sticky=W + E)
        self.set_form_values(
            sheet_button_text="Select",
            message_button_text="Select",
            error_text=""
        )
        self.base.mainloop()

    def set_form_values(self, sheet_button_text=None, message_button_text=None, account_selection=None, error_text=None):
        if sheet_button_text is not None:
            self.select_sheet_button_text.set(sheet_button_text)
        if message_button_text is not None:
            self.select_message_button_text.set(message_button_text)
        if error_text is not None:
            self.error_message_text.set(error_text)
        # Reset var and delete all old options
        self.account_menu['menu'].delete(0, 'end')
        # Insert list of new options (tk._setit hooks them up to var)
        account_list = list(get_stored_accounts().keys())
        account_list.append("Add New/Update Existing")
        for choice in account_list:
            self.account_menu['menu'].add_command(label=choice,
                                                  command=tkinter._setit(self.account_menu_selection, choice))
        if account_selection is not None:
            self.account_menu_selection.set(account_selection)
            self.selected_account = account_selection
        elif self.selected_account == "" or self.selected_account == None:
            self.account_menu_selection.set(account_list[0])
            self.selected_account = account_list[0]


    def select_sheet(self):
        filepath = tkinter.filedialog.askopenfile().name
        if Path(filepath).suffix != ".csv":
            self.set_form_values(
                error_text="Contact Sheet Must be a CSV Document"
            )
            return
        self.selected_sheet = filepath
        self.set_form_values(
            sheet_button_text=f"[{Path(self.selected_sheet).name}]",
            error_text=""
        )

    def select_message(self):
        filepath = tkinter.filedialog.askopenfile().name
        if Path(filepath).suffix != ".html":
            self.set_form_values(
                error_text="Message Template Must be an HTML document"
            )
            return
        self.selected_message = filepath
        self.set_form_values(
            message_button_text=f"[{Path(self.selected_message).name}]",
            error_text=""
        )

    def validate_sheet(self):
        if self.selected_sheet == "" or self.selected_sheet is None:
            self.set_form_values(
                error_text="Please select a valid CSV file to validate."
            )
            return
        validate_csv_file(self.selected_sheet)

    def option_menu_click(self, *args):
        if self.account_menu_selection.get() == "Add New/Update Existing":
            AccountMenu(self.base)
        else:
            self.set_form_values(
                account_selection= self.account_menu_selection.get()
           )

    def send_emails(self):
        if self.selected_sheet == "" or self.selected_sheet == None and self.selected_message == "" or self.selected_message == None:
            self.set_form_values(
                error_text="Please select a contact sheet and a message sheet."
            )
            return
        if self.selected_sheet == "" or self.selected_sheet == None:
            self.set_form_values(
                error_text="Please select a contact sheet."
            )
            return
        if self.selected_message == "" or self.selected_message == None:
            self.set_form_values(
                error_text="Please select a message template."
            )
            return

        start_send_emails(self.selected_sheet, self.selected_message, self.selected_account, get_stored_accounts()[self.selected_account])