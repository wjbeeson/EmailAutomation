import json
import smtplib, ssl, csv
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart

from email_validator import validate_email, EmailNotValidError
import html2text
import re


def get_validated_email(email):
    try:
        # validate and get info
        validation_info = validate_email(email)

        # replace with normalized form
        email = validation_info.email
        return email
    except EmailNotValidError as e:
        # email is not valid, exception message is human-readable

        print(f"ERROR {email} [{str(e)}] Skipping to next entry...")
        return None


def send_email(server, email, subject, message, sender):
    email = "wbeeson@autelrobotics.com"
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

        #  Create an Email Image object
        fp = open("assets/" + image, 'rb')
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


def validate_message(column_names):
    html_template = open("assets/email_message.html").read()

    #  Check to make sure user submitted the correct tag pairs
    tags = [tag.lower().strip() for tag in re.findall(r"\{(.*?)\}", html_template)]
    if sorted(tags) != sorted(column_names):
        while True:
            print("Detected tags are different than the tags defined in the CSV file.")
            print(f"CSV:  [{sorted(column_names)}]")
            print(f"HTML: [{sorted(tags)}]")
            response = input("Continue? (y/n)")
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


def start_server():
    with open('keys/email.json') as f:
        file = json.load(f)
        sender = file["username"]
        password = file["password"]

    #  Start the email server
    context = ssl.create_default_context()
    server = smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context)
    server.login(sender, password)
    return server, sender


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
smtp_server, sender = start_server()

#  Load the contact information
csv_file = open('assets/list.csv', "r", encoding="UTF-8")
csv_reader = csv.reader(csv_file, delimiter=',')

#  Get the column names
email_index, column_names = get_column_names(csv_reader)
template = validate_message(column_names)

#  Get the subject
subject = input("Please enter the subject line: ")

#  Send the Emails
for i, rows in enumerate(csv_reader):
    column_values = get_row_values(rows)
    email = get_validated_email(column_values.pop(email_index))
    if email is None:
        continue
    get_customized_message = customize_message(template, column_names, column_values)
    send_email(
        server=smtp_server,
        email=email,
        subject=subject,
        message=get_customized_message,
        sender=sender
    )
    print(f"Message Sent To: {email} [Row {i}]")

with open('assets/list.csv') as f:
    row_count = sum(1 for line in f) - 1
print(f"Done: Send Emails to all {row_count} Contacts")
