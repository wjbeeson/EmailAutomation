import json
import smtplib, ssl, csv
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart


def send_email(subject, text_body, html_body, sender, recipients: list, password, server):
    # Declare message root
    msg_root = MIMEMultipart("related")
    msg_root['Subject'] = subject
    msg_root['From'] = sender
    msg_root['To'] = ', '.join(recipients)

    # Assign message alternatives
    msg_alternative = MIMEMultipart('alternative')
    msg_root.attach(msg_alternative)
    msg_alternative.attach(text_body)
    msg_alternative.attach(html_body)

    # Add image
    fp = open('assets/logo.png', 'rb')
    msg_image = MIMEImage(fp.read())
    fp.close()

    # Define the image's ID as referenced in HTML text
    msg_image.add_header('Content-ID', '<logo>')
    msg_root.attach(msg_image)

    # Send message
    server.sendmail(sender, recipients, msg_root.as_string())


def send_old_greeting_email(name, product, recipient, server):
    # Step 1: Create the plain text variant
    text = ("Dear {name},\n\nI am writing to follow up on an information query you made some time ago about our "
            "{product}. I am checking in to make sure that you have found a solution that you are satisfied with. \n\n"
            "What solution did you end up going with?\n\nSincerely,\n\nWill Beeson | Sales Representative\n"
            "Direct +1 (425)780-8024\nwww.autelrobotics.com")

    # Step 2: Create the plain html variant
    html = """\
    <html>
      <body>
        <p>Dear {name},</p>
        <p>I am writing to follow up on an information query you made some time ago about our {product}. I am checking in 
            to make sure that you have found a solution that you are satisfied with.</p>
        <p>What solution did you end up going with?</p>
        <p>Sincerely,</p>
        <p><b>Will Beeson | Sales Representative</b><br>Direct +1 (425)780-8024<br><img src="cid:logo"><br>www.autelrobotics.com</p>
      </body>
    </html>
    """

    # Step 3: Set email information
    text = text.format(name=name, product=product)
    html = html.format(name=name, product=product)
    set_html_body = MIMEText(html, "html")
    set_text_body = MIMEText(text, "plain")

    subject = "Drone Information Follow-up"
    send_email(subject, set_text_body, set_html_body, sender, recipient, password, server)


def send_livedeck_promotion_email(name, dealer, recipient, server):
    # Step 1: Create the plain text variant
    text = (
        "Dear {name},\n\n"

        "I hope this message finds you well. I'm excited to share an exclusive offer with {dealer} as a "
        "valued partner of Autel Robotics. For a limited time only, we're thrilled to provide you with a special "
        "offer: with every purchase of an Autel Robotics EVO II 640T V3 series drone, you'll receive a "
        "complimentary LiveDeck 2.\n\n"

        "The LiveDeck 2 seamlessly integrates with the EVO II 640T V3 drone to deliver high-quality, real-time video "
        "streaming capabilities without requiring any wifi or cellular activity to use. \n\n"

        "This offer is available until May 1st. Contact us today to place your order and take advantage of this "
        "special deal! This offer cannot be combined with any other offer. \n\n"


        "Thank you for your continued partnership with Autel Robotics. We're committed to supporting you and helping "
        "you succeed in meeting your customers' needs and expectations.\n\n"

        "Will Beeson | Sales Representative\n"
        "www.autelrobotics.com")

    # Step 2: Create the html variant
    html = """\
    <html>
      <body>
        <p>Dear {name},</p>

        <p>I hope this message finds you well. I'm excited to share an exclusive offer with {dealer} as a valued partner 
        of Autel Robotics. For a limited time only, we're thrilled to provide you with a special offer: with every 
        purchase of an Autel Robotics EVO II 640T V3 series drone, you'll receive a complimentary LiveDeck 2.</p>

        <p>The LiveDeck 2 seamlessly integrates with the EVO II 640T V3 drone to deliver high-quality, real-time video streaming 
        capabilities without requiring any wifi or cellular activity to use. </p>

        <p>This offer is available until May 1st. Contact us today to place your order and take advantage of this 
        special deal! This offer cannot be combined with any other offer. </p>

        <p>Thank you for your continued partnership with Autel Robotics. We're committed to supporting you and helping 
        you succeed in meeting your customers' needs and expectations.</p>
        
        <p><a href="https://drive.google.com/drive/folders/1m1GO5SwB82gZkI4CVvZ1b8aGZRSuR2HX">LiveDeck 2 Information</a><br>
        <a href="https://drive.google.com/drive/folders/1-CxA0Q-Ch2s4YMNu5sUiVLTWBqEhc5K7">EVO II 640T Information</a>
        </p>
        
        <p><img src="cid:logo" width: 30%;><b><br>
        Will Beeson | Sales Representative</b><br>
      </body>
    </html>
    """

    # Step 3: Set email information
    text = text.format(name=name, dealer=dealer)
    html = html.format(name=name, dealer=dealer)
    set_html_body = MIMEText(html, "html")
    set_text_body = MIMEText(text, "plain")

    subject = "Autel Offer: Free LiveDeck 2 with Each EVO II Purchase!"
    send_email(subject, set_text_body, set_html_body, sender, recipient, password, server)


with open('keys/will_email.json') as f:
    file = json.load(f)
    sender = file["username"]
    password = file["password"]

context = ssl.create_default_context()
with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp_server:
    smtp_server.login(sender, password)
    with open('assets/dealer_list.csv', "r", encoding="UTF-8") as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=',')
        for i, row in enumerate(csv_reader):
            if i == 0: continue
            dealer = row[0]
            name = row[1]
            email = row[2]
            send_livedeck_promotion_email(name=name, dealer=dealer, recipient=[email], server=smtp_server)
            print(f"Message Sent To: {email} [Row {i}]")

with open('assets/dealer_list.csv') as f:
    row_count = sum(1 for line in f) - 1
print(f"Done: Send Emails to all {row_count} Contacts")
