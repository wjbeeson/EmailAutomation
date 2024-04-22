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
import threading


def get_validated_email(email, i, errors, increment, progress_bar, base, submissions):
    try:
        # validate and get info
        validation_info = validate_email(email)

        # replace with normalized form
        email = validation_info.normalized
    except EmailNotValidError as e:
        errors[i] = email + " " + f"[{e}]"
    submissions.append(i)
    progress_bar.step(increment)


def validate_csv_file(csv_filename, progress_bar, base):
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
    increment = 99.9 / lines

    #  Get Emails
    errors = {}
    submissions = []

    #  Validate emails
    for i, rows in enumerate(csv_reader):
        base.update_idletasks()
        base.update()
        column_values = get_row_values(rows)
        raw_email = column_values.pop(email_index)
        t = threading.Thread(target=get_validated_email,
                             args=(raw_email, i, errors, increment, progress_bar, base, submissions))
        t.daemon = True
        t.start()

    #  Wait for submissions to be done
    while len(submissions) < lines:
        base.update_idletasks()
        base.update()

    #  Cap the displayed error count
    total_error_count = len(errors)
    max_error_display = 10
    sorted_errors_keys = []
    for i, key in enumerate(sorted(errors)):
        if i < max_error_display:
            sorted_errors_keys.append(key)

    #  Sort error messages
    sorted_errors = {}
    for i in sorted_errors_keys:
        sorted_errors[i] = errors[i]

    #  Print out error messages
    error_message = ""
    if len(errors) > 0:
        error_message += f"\nShowing {min(total_error_count, max_error_display)} of {total_error_count} errors: \n"
        for error in sorted_errors:
            error_message += f"\nLine {error + 2}: \t{sorted_errors[error]}\n"
        messagebox.showerror("Errors", error_message)
    else:
        messagebox.showinfo("Woohoo!", "No detected errors")
    return email_index, column_names
