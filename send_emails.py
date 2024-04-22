import queue
import smtplib
import ssl
import csv
import threading
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
import time
import pandas as pd
import html2text
from tkinter import messagebox
from validate_images import *


def ensure_csv_has_email_field(csv_filename):
    csv_file = open(csv_filename, "r", encoding="UTF-8")
    csv_reader = csv.reader(csv_file, delimiter=',')

    #  Check the first line for the required email field
    column_names = get_row_values(next(csv_reader))
    if "email" not in column_names:
        error_message = (f"Your top line header labels of the CSV file must contain the field 'email'.\n"
                         f"Detected fields: {column_names}")
        messagebox.showinfo("Errors", error_message)
        raise Exception("Include the field 'Email' in the CSV file and rerun")

    #  Remove the email field as it is hardcoded for different functionality than the rest of them
    email_index = column_names.index("email")
    column_names.pop(email_index)
    return email_index, column_names, csv_reader


def open_validate_message(column_names, message_filename):
    template = open(message_filename).read()

    #  Check to make sure user submitted the correct tag pairs
    tags = [tag.lower().strip() for tag in re.findall(r"\{(.*?)}", template)]
    if sorted(tags) != sorted(column_names):
        while True:
            response = messagebox.askquestion("Continue?",
                                              "HTML tags are different than the headers defined in the CSV file.\n"
                                              f"CSV:    {sorted(column_names)}\n"
                                              f"HTML: {sorted(tags)}\n\nContinue?")
            if response.lower().__contains__("y"):
                break
            if response.lower().__contains__("n"):
                raise Exception("User stopped the program. Revise tags and re-run.")

    return template


def customize_message(template, column_names, column_values):
    for i, tag in enumerate(column_names):
        template = template.replace("{" + tag + "}", column_values[i].title())
    return template


def get_row_values(rows):
    values = []
    for i in range(len(rows)):
        values.append(rows[i].lower().strip())
    return values


def start_server(email, password):
    #  Start the email server
    context = ssl.create_default_context()
    server = smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context)
    # server.set_debuglevel(1)
    server.login(email, password)
    return server


def get_column_names(csv_reader):
    #  Check the first line for the required email field
    column_names = get_row_values(next(csv_reader))
    if "email" not in column_names:
        raise Exception(f"Your top line header labels of the CSV file must contain the field 'email'. "
                        f"Detected fields: {column_names}")

    #  Remove the email field as it is hardcoded for different functionality than the rest of them
    email_index = column_names.index("email")
    column_names.pop(email_index)
    return email_index, column_names


def format_message_root(subject, username, template, image_paths, message_filepath, email):
    # Declare message root
    msg_root = MIMEMultipart("related")
    msg_root['Subject'] = subject
    msg_root['From'] = username
    msg_root["To"] = email
    msg_alternative = MIMEMultipart('alternative')
    msg_root.attach(msg_alternative)

    return msg_root, msg_alternative


def add_images_to_root(msg_root, message_filepath, image_paths, message):
    #  Add images
    images = get_image_filenames(message_filepath)
    for i, image in enumerate(images):

        #  Set the ID of the image in the text
        message = message.replace(image, "cid:" + str(i), 1)
        fp = open(image_paths[i], 'rb')
        msg_image = MIMEImage(fp.read())
        fp.close()

        # Define the image's ID as referenced in HTML text
        msg_image.add_header('Content-ID', f'<{i}>')
        msg_root.attach(msg_image)
    return message


def assemble_payload(subject, sender, recipient, message, image_paths, message_filepath):
    #  Format the message root
    msg_root, msg_alternative = format_message_root(subject, sender, message, image_paths, message_filepath, recipient)

    #  Add images to root and add content-ids to message
    message = add_images_to_root(msg_root, message_filepath, image_paths, message)

    #  Add plain text version to Email
    text_version = html2text.html2text(message)
    text_final = MIMEText(text_version, "plain")
    msg_alternative.attach(text_final)

    #  Add HTML version to Email
    html_final = MIMEText(message, "html")
    msg_alternative.attach(html_final)

    return msg_root, recipient


def assemble_message(rows, email_index, template, column_names, subject, sender, image_paths, message_filepath):
    #  Pull out the email field from the column values
    column_values = get_row_values(rows)
    recipient = column_values.pop(email_index)

    #  Format the template with column info
    message = customize_message(template, column_names, column_values)

    msg_root, recipient = assemble_payload(subject, sender, recipient, message, image_paths, message_filepath)
    message = Message(msg_root, recipient)
    return message


def email_manager(csv_filepath, message_filepath, sender, password, subject, image_paths, progress_bar, base):
    #  Calculate Increment for progress bar
    lines = len(pd.read_csv(csv_filepath))
    total_progress_bar_size = 99.9
    increment = total_progress_bar_size / lines

    #  Get column information
    email_index, column_names, csv_reader = ensure_csv_has_email_field(csv_filepath)

    #  Open message and validate that the message fields match the csv columns
    template = open_validate_message(column_names, message_filepath)

    #  Define Server Count
    server_count = 5

    #  Create the message queue
    message_queues = []
    for i in range(server_count):
        message_queues.append([])

    #  Assemble the message queues
    for i, rows in enumerate(csv_reader):
        message = assemble_message(rows, email_index, template, column_names, subject, sender, image_paths,
                                   message_filepath)
        message_queues[i % server_count].append(message)

    #  Send Emails
    for i in range(server_count):
        t = threading.Thread(target=send_emails,
                             args=(sender, password, increment, progress_bar, base, message_queues[i], i))
        t.daemon = True
        t.start()

    #  Wait for all emails to send
    while True:
        keep_waiting = False
        for msg_queue in message_queues:
            if len(msg_queue) > 0:
                keep_waiting = True
        if not keep_waiting:
            break
        base.update_idletasks()
        base.update()

    final_message = f"Sent Emails to {lines} contacts\n"
    messagebox.showinfo("Done", final_message)


def send_emails(sender, password, increment, progress_bar, base, message_queue: list, i):
    #  Start the server and return if failed
    try:
        server = start_server(sender, password)
    except Exception as e:
        if i == 0:
            messagebox.showinfo("Error", str(e))
        message_queue.clear()
        return

    #  Send the messages and ignore any errors
    while len(message_queue) > 0:
        message = message_queue.pop()
        try:
            server.sendmail(sender, [message.recipient], message.msg_root.as_string())
        except:
            pass

        #  Update the progress bar
        progress_bar.step(increment)
        base.update_idletasks()
        base.update()

    #  server.quit()


class Message():
    def __init__(self, msg_root, recipient):
        self.msg_root = msg_root
        self.recipient = recipient
