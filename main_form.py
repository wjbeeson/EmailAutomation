import tkinter.filedialog
from account_form import *
from account_management import *
from send_emails import *
from tkinter import Label
from tkinter import *
from pathlib import Path
from functools import partial
from validate_images import *
from tkinter import ttk
from validate_sheet import validate_csv_file


class MainMenu:
    def __init__(self, root):
        for ele in root.winfo_children():
            ele.destroy()
        self.base = root
        self.selected_sheet = ""
        self.selected_message = ""
        self.selected_account = ""
        self.color = "Green"

        self.base.geometry('450x400')  # w x h
        self.base.title("Email Bot")
        current_row = 0

        self.main_frame = Frame(self.base, width=450, height=400)
        self.main_frame.pack()

        # Select the Subject
        Label(self.main_frame, text="Subject: ", font=("bold", 10)).grid(row=current_row, column=0, columnspan=1,
                                                                         sticky=W)
        self.subject = StringVar()
        Entry(self.main_frame, textvariable=self.subject, font=("bold", 10)).grid(row=current_row, column=1,
                                                                                  columnspan=1,
                                                                                  sticky=E + W)
        current_row += 1

        #  Dropdown menu to change account
        Label(self.main_frame, text="Account: ", font=("bold", 10)).grid(row=current_row, column=0, columnspan=1,
                                                                         sticky=W)
        self.account_menu_selection = tkinter.StringVar()
        self.account_menu_selection.trace("w", self.option_menu_click)  # Trace to call function whenever changed
        self.account_menu = OptionMenu(self.main_frame, self.account_menu_selection, *[""])
        self.account_menu.grid(row=current_row, column=1, sticky=EW, columnspan=1)
        current_row += 1

        # Buttons to select and validate a CSV sheet
        Label(self.main_frame, text="Contact Sheet: ", font=("bold", 10)).grid(row=current_row, column=0, columnspan=1,
                                                                               sticky=W)
        self.select_sheet_button_text = tkinter.StringVar()
        self.select_sheet_button = Button(self.main_frame, textvariable=self.select_sheet_button_text, fg='black',
                                          command=self.select_sheet)
        self.select_sheet_button.grid(row=current_row, column=1, columnspan=1, sticky=W + E)
        Button(self.main_frame, text='Validate Emails', bg=self.color, fg='white', command=self.validate_sheet).grid(
            row=current_row,
            column=2,
            columnspan=1,
            sticky=W + E)
        current_row += 1

        # Select message button
        Label(self.main_frame, text="Message Template: ", font=("bold", 10)).grid(row=current_row, column=0, sticky=W,
                                                                                  columnspan=1)
        self.select_message_button_text = tkinter.StringVar()
        self.select_message_button = Button(self.main_frame, textvariable=self.select_message_button_text, fg='black',
                                            command=self.select_message)
        self.select_message_button.grid(
            row=current_row,
            column=1,
            columnspan=1,
            sticky=W + E)
        current_row += 1

        # Add image subframe for later adding image validation if needed
        self.image_subframe = Frame(self.main_frame, width=self.main_frame.winfo_width(), padx=20)
        self.image_subframe.grid(
            row=current_row,
            column=0,
            columnspan=4,
            sticky=W + E)
        self.image_selector_text_map = {}  # Used to store references to the StringVars on the image selector buttons
        self.image_selector_map = {}  # Used to store the actual paths for the selected images
        current_row += 1

        # Add send emails button
        self.send_emails_button = Button(self.main_frame, text='Send Emails', bg=self.color, fg='white',
                                         command=self.send_emails)
        self.send_emails_button.grid(
            row=current_row,
            column=0,
            columnspan=4,
            sticky=W + E)
        current_row += 1

        # Add configurable error text to avoid excessive pop-up windows
        self.error_message_text = tkinter.StringVar()
        self.error_message = Label(self.main_frame, textvariable=self.error_message_text, font=("bold", 10),
                                   foreground="red")
        current_row += 1

        self.error_message.grid(row=current_row, column=0, columnspan=4, sticky=W + E)
        self.set_form_values(
            sheet_button_text="Select",
            message_button_text="Select",
            error_text=""
        )
        current_row += 1

        # add progress bar subframe
        self.progress_bar = None
        self.progress_subframe = Frame(self.main_frame, width=self.main_frame.winfo_width(), padx=0)
        self.progress_subframe.columnconfigure(0, weight=1)
        self.progress_subframe.grid(
            row=current_row,
            column=0,
            columnspan=4,
            sticky=W + E)
        current_row += 1

        self.base.mainloop()

    def add_progress_bar(self):
        self.remove_progress_bar()
        self.progress_bar = ttk.Progressbar(self.progress_subframe)
        self.progress_bar.grid(row=0, column=0, sticky=W + E)

    def remove_progress_bar(self):
        for ele in self.progress_subframe.winfo_children():
            ele.destroy()

    def add_image_selectors(self, frame: Frame, defined_image_names: list):
        # Clear all previous images from the selector panel

        self.image_selector_text_map = {}

        current_row = 0
        for i, name in enumerate(defined_image_names):
            # Select message button
            Label(frame, text=f"Select {name}: ", font=("bold", 10)).grid(row=current_row, column=0, columnspan=1,
                                                                          sticky=W)
            # Need to use an ID because can't assume image name is unique
            selector_id = i
            self.image_selector_text_map[selector_id] = StringVar()
            self.image_selector_text_map[selector_id].set(value="Select")

            action_with_arg = partial(self.select_image, selector_id)
            image_selector = Button(frame, textvariable=self.image_selector_text_map[selector_id], fg='black',
                                    command=action_with_arg)
            image_selector.grid(row=current_row, column=1, columnspan=1, sticky=W + E)
            current_row += 1

    def select_image(self, selector_id):
        filepath = tkinter.filedialog.askopenfile().name
        if Path(filepath).suffix != ".jpg" and Path(filepath).suffix != ".png":
            self.set_form_values(
                error_text="Allowed image types include .jpg or .png."
            )
            return
        self.image_selector_map[selector_id] = filepath
        self.set_form_values(
            update_image_selectors=True,
            error_text=""
        )

    def set_form_values(self, sheet_button_text=None, message_button_text=None, account_selection=None,
                        error_text=None, update_image_selectors=False):
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
        if update_image_selectors:
            for selector in list(self.image_selector_map.keys()):
                self.image_selector_text_map[selector].set(Path(self.image_selector_map[selector]).name)

    def select_sheet(self):
        filepath = tkinter.filedialog.askopenfile().name
        if Path(filepath).suffix != ".csv":
            self.set_form_values(
                error_text="Contact Sheet Must be a CSV Document"
            )
            return
        self.selected_sheet = filepath

        self.validate_sheet()

        self.set_form_values(
            sheet_button_text=f"[{Path(self.selected_sheet).name}]",
            error_text=""
        )

    def validate_sheet(self):
        if self.selected_sheet == "" or self.selected_sheet is None:
            self.set_form_values(
                error_text="Please select a valid CSV file."
            )
            return
        self.toggle_base_state('disable')
        self.add_progress_bar()
        validate_csv_file(self.selected_sheet, self.progress_bar, self.base)
        self.remove_progress_bar()
        self.toggle_base_state()

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

        image_filepaths = get_image_filenames(filepath)
        self.add_image_selectors(self.image_subframe, image_filepaths)

    def option_menu_click(self, *args):
        if self.account_menu_selection.get() == "Add New/Update Existing":
            AccountMenu(self.base)
        else:
            self.set_form_values(
                account_selection=self.account_menu_selection.get()
            )

    def toggle_base_state(self, state='normal'):
        for child in self.main_frame.winfo_children():
            try:
                child['state'] = state
            except:
                pass
        for child in self.image_subframe.winfo_children():
            try:
                child['state'] = state
            except:
                pass

    def send_emails(self):
        if self.subject.get().replace(" ", "") == "" or self.subject == None:
            self.set_form_values(
                error_text="Please enter a subject."
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
        if sorted(list(self.image_selector_map.keys())) != sorted(list(self.image_selector_text_map.keys())):
            self.set_form_values(
                error_text="Please select the paths for all images in the message."
            )
            return
        self.set_form_values(
            error_text=""
        )
        # Need to sort image paths to make them line up
        self.toggle_base_state('disabled')
        self.add_progress_bar()
        sorted_image_paths = []
        for selector in sorted(self.image_selector_map):
            sorted_image_paths.append(self.image_selector_map[selector])

        email_manager(
            csv_filepath=self.selected_sheet,
            message_filepath=self.selected_message,
            sender=self.selected_account,
            password=get_stored_accounts()[self.selected_account],
            subject=self.subject.get(),
            image_paths=sorted_image_paths,
            progress_bar=self.progress_bar,
            base=self.base)
        self.remove_progress_bar()
        self.toggle_base_state()
