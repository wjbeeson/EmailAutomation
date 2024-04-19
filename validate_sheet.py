import json
import math
import smtplib
import ssl
import csv
import time
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email_validator import validate_email, EmailNotValidError
import html2text
import re
from tkinter import filedialog
from tkinter import simpledialog
from tkinter import messagebox
from send_emails import get_row_values
import pandas as pd


def get_validated_email(email):
    try:
        # validate and get info
        validation_info = validate_email(email)

        # replace with normalized form
        email = validation_info.normalized
        return email, None
    except EmailNotValidError as e:
        # email is not valid, return error message
        return None, str(e)


def validate_csv_file(csv_filename, progress_bar_progress, base):
    errors = {}
    csv_file = open(csv_filename, "r", encoding="UTF-8")
    csv_reader = csv.reader(csv_file, delimiter=',')

    #  Check the first line for the required email field
    column_names = get_row_values(next(csv_reader))
    if "email" not in column_names:
        error_message = (f"Your top line header labels of the CSV file must contain the field 'email'.\n"
                         f"Detected fields: {column_names}")
        messagebox.showerror("Errors", error_message)

    #  Remove the email field as it is hardcoded for different functionality than the rest of them
    email_index = column_names.index("email")
    column_names.pop(email_index)

    #  Calculate Increment for progress bar
    lines = len(pd.read_csv(csv_filename))
    total_progress_bar_size = 99.9
    increment = total_progress_bar_size / lines
    current_progress = 0

    #  Validate emails
    for i, rows in enumerate(csv_reader):
        column_values = get_row_values(rows)
        raw_email = column_values.pop(email_index)
        email, error = get_validated_email(raw_email)
        if error is not None:
            errors[i] = raw_email + " " + f"[{error}]"
        current_progress += increment
        progress_bar_progress.set(current_progress)
        base.update_idletasks()
        base.update()
    error_message = ""
    if len(errors) > 0:
        error_message += f"\nErrors: "
        for error in errors:
            error_message += f"\nLine {error + 1}: \t{errors[error]}\n"
        messagebox.showerror("Errors", error_message)
    else:
        messagebox.showinfo("Woohoo!", "No detected errors")
    return email_index, column_names
