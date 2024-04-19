import json
import smtplib
import ssl
import csv
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email_validator import validate_email, EmailNotValidError
import html2text
import re
from tkinter import filedialog
from tkinter import simpledialog
from tkinter import messagebox



def validate_images(message):
    images = re.findall('<img [^>]*src="([^"]+)', message)
    image_filenames = []
    for i, image in enumerate(images):
        #  Get the path for the image in case they used a local reference earlier
        messagebox.showinfo(f"Load Image {i + 1}", f"Please select the image {image}")
        image_filename = filedialog.askopenfilename()
        image_filenames.append(image_filename)
    return image_filenames


def ensure_csv_has_email_field(csv_filename):
    errors = {}
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
    return email_index, column_names


def send_email(server, email, subject, message, sender, image_paths):
    # Declare message root
    msg_root = MIMEMultipart("related")
    msg_root['Subject'] = subject
    msg_root['From'] = sender
    msg_root['To'] = email

    # Assign message alternatives
    msg_alternative = MIMEMultipart('alternative')
    msg_root.attach(msg_alternative)

    #  Add images to the HTML version
    images = re.findall('<img [^>]*src="([^"]+)', message)
    for i, image in enumerate(images):
        #  Set the ID of the image in the text
        message = message.replace(image, "cid:" + str(i))

        fp = open(image_paths[i], 'rb')
        msg_image = MIMEImage(fp.read())
        fp.close()

        # Define the image's ID as referenced in HTML text
        msg_image.add_header('Content-ID', f'<{i}>')
        msg_root.attach(msg_image)

    #  Add plain text version to Email
    text_version = html2text.html2text(message)
    text_final = MIMEText(text_version, "plain")
    msg_alternative.attach(text_final)

    #  Add HTML version to Email
    html_final = MIMEText(message, "html")
    msg_alternative.attach(html_final)

    server.sendmail(sender, [email], msg_root.as_string())


def validate_message(column_names, message_filename):
    html_template = open(message_filename).read()

    #  Check to make sure user submitted the correct tag pairs
    tags = [tag.lower().strip() for tag in re.findall(r"\{(.*?)}", html_template)]
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

    return html_template


def customize_message(template, column_names, column_values):
    for i, tag in enumerate(column_names):
        template = template.replace("{" + tag + "}", column_values[i].title())
    return template


def get_row_values(rows):
    values = []
    for i in range(len(rows)):
        values.append(rows[i].lower().strip())
    return values


def start_server(username, password):

    #  Start the email server
    context = ssl.create_default_context()
    server = smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context)
    server.login(username, password)
    return server, username


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


#  Start the server
def start_send_emails(csv_filepath, message_filepath, username, password):
    smtp_server, sender = start_server(username, password)



    #  Get the column names
    email_index, column_names = ensure_csv_has_email_field(csv_filepath)

    csv_file = open(csv_filepath, "r", encoding="UTF-8")
    csv_reader = csv.reader(csv_file, delimiter=',')
    next(csv_reader)

    template = validate_message(column_names, message_filepath)

    #  Validate CSV file emails

    #  Validate the image paths used for the html version
    image_paths = validate_images(template)

    #  Get the subject
    while True:
        subject = simpledialog.askstring("Subject", "Please enter the subject line: ")
        response = messagebox.askquestion(f"Confirm", f"Confirm Subject: \n\"{subject}\"")
        if response.lower().__contains__("y"):
            break

    #  Send the Emails
    emails_sent = 0
    for i, rows in enumerate(csv_reader):
        column_values = get_row_values(rows)
        email = column_values.pop(email_index)

        get_customized_message = customize_message(template, column_names, column_values)
        send_email(
            server=smtp_server,
            email=email,
            subject=subject,
            message=get_customized_message,
            sender=sender,
            image_paths=image_paths
        )
        print(f"Message Sent To: {email} [Row {i}]")
        emails_sent += 1

    final_message = f"Sent Emails to {emails_sent} contacts\n"
    messagebox.showinfo("Done", final_message)


