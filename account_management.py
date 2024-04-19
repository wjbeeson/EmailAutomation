import os
import glob
import json
import smtplib
import uuid
from tkinter import messagebox

KEY_DIRECTORY = "C:\\keys"
def get_stored_accounts():
    if not os.path.isdir(KEY_DIRECTORY):
        os.mkdir(KEY_DIRECTORY)
    keys = glob.glob(f"{KEY_DIRECTORY}/*.json")

    account_dict = {}
    for file in keys:
        account_info = json.load(open(file))
        account_dict[account_info["username"]] = account_info["password"]
        pass
    return account_dict

def test_valid_credentials(username, password):
    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    try:
        server.login(username, password)
        server.quit()
        return True
    except Exception:
        server.quit()
        return False

def add_account(username, password):
    account_dict = {"username": username, "password": password}
    existing_accounts = get_stored_accounts()
    keys = glob.glob(f"{KEY_DIRECTORY}/*.json")
    replace_file = None
    for i, account in enumerate(list(existing_accounts.keys())):
        if username == account:
            replace_file = keys[i]
            break
    new_keyname = f"{KEY_DIRECTORY}\\{uuid.uuid4()}.json"
    if replace_file is not None:
        new_keyname = replace_file
    with open(new_keyname, "w") as file:
        json.dump(account_dict, file)

    if replace_file:
        messagebox.showinfo("Updated Credentials", f"Updated Credentials for {username}")
        return
    messagebox.showinfo("Added Credentials", f"Added Credentials for {username}")


