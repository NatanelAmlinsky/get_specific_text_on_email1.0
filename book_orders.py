import base_page
from base_page import OutlookAccount
import json


class OrderBooksForm:
    json_file = open('configuration.json', 'r', encoding='utf-8')  # open the JSON file
    data = json.load(json_file)  # loading the JSON content to the variable "data"
    account_name = data["account name"]
    outlook_account = OutlookAccount(account_name)


    if outlook_account.login():
        # set up the sender name and email to filter by
        senders_name = data["sender name"]
        senders_email = data["senders email"]
        count = 0
        forms_txt_subject = data["forms txt subject"]

        # loop through the emails in the inbox folder
        for message in outlook_account.inbox_folder.Items:
            if message.Class == 43:
                if message.SenderEmailType == "EX":
                    for email in senders_email:
                        if message.Sender.GetExchangeUser().PrimarySmtpAddress == email:
                            for subject in forms_txt_subject:
                                if subject in message.Subject.lower():
                                    # get the email content
                                    parts = base_page.OutlookAccount.get_email_content(message)
                                    count += 1
                                    break  # exit the inner loop
                else:
                    for email in senders_email:
                        if message.SenderEmailAddress == email:
                            for subject in forms_txt_subject:
                                if subject in message.Subject.lower():
                                    # get the email content
                                    parts = base_page.OutlookAccount.get_email_content(message)
                                    count += 1
                                    break  # exit the inner loop

        print("Found {} emails from '{}'".format(count, senders_email))

