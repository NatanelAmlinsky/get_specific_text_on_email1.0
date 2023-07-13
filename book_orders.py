import base_page
from base_page import OutlookAccount
import json


class OrderBooksForm:
    json_file = open('configuration.json', 'r', encoding='utf-8')
    data = json.load(json_file)
    account_name = ""
    senders_email = ""
    count = 0

    for account_name_txt in data["account name"]:
        if account_name_txt in data["account name"]:
            account_name = account_name_txt
            outlook_account = OutlookAccount(account_name)

            if outlook_account.login():
                senders_email = data["senders email"]
                count = 0
                forms_txt_subject = data["forms txt subject"]

                for message in outlook_account.inbox_folder.Items:
                    if message.Class == 43:
                        if message.SenderEmailType == "EX":
                            for email in senders_email:
                                if email in message.SenderEmailAddress:
                                    for unusual_email in data["Unusual emails"]:
                                        if email in unusual_email:
                                            parts = outlook_account.get_email_content(message)
                                            count += 1
                                            break
                                    for subject in forms_txt_subject:
                                        if subject in message.Subject.lower():
                                            parts = outlook_account.get_email_content(message)
                                            count += 1
                                            break
                        else:
                            for email in senders_email:
                                if email in message.SenderEmailAddress:
                                    for unusual_email in data["Unusual emails"]:
                                        if email in unusual_email:
                                            parts = outlook_account.get_email_content(message)
                                            count += 1
                                            break
                                    for subject in forms_txt_subject:
                                        if subject in message.Subject.lower():
                                            parts = outlook_account.get_email_content(message)
                                            count += 1
                                            break

    print("Found {} emails from '{}'".format(count, senders_email), account_name)
