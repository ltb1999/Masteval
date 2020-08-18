import imaplib
import email
from email.header import decode_header
from read_from_excel import add_dates
import webbrowser
import os
import datetime


class Trainee:
    def __init__(self, type, firstname, name2, date):
        self.type = type
        self.firstname = firstname
        self.name2 = name2
        self.date = date


def read_email(number):
    # account credentials
    username = "darlinghuynh2912@gmail.com"
    password = "@Ltb1999"

    # create an IMAP4 class with SSL
    imap = imaplib.IMAP4_SSL("imap.gmail.com")
    # authenticate
    imap.login(username, password)

    status, messages = imap.select("INBOX")
    # number of top emails to fetch
    N = number
    list_of_trainee = []
    # total number of emails
    messages = int(messages[0])

    for i in range(messages, messages - N, -1):
        # fetch the email message by ID
        res, msg = imap.fetch(str(i), "(RFC822)")
        for response in msg:
            if isinstance(response, tuple):
                # parse a bytes email into a message object
                msg = email.message_from_bytes(response[1])
                # decode the email subject
                subject = decode_header(msg["Subject"])[0][0]
                if isinstance(subject, bytes):
                    # if it's a bytes, decode to str
                    subject = subject.decode()
                # email sender
                from_ = msg.get("From")
                # print("Subject:", subject)
                # Check the type of the email
                type_of_email = ["Host Trainer Mid-Point", "Host Trainer Final", "Trainee/Intern Initial",
                                 "Trainee/Intern Mid-Point", "Trainee/Intern Final"]
                for index in range(0, len(type_of_email)):
                    if subject.find(type_of_email[index]) != -1:
                        real_type = index
                        break
                if msg.is_multipart():
                    # iterate over email parts
                    for part in msg.walk():
                        # extract content type of email
                        content_type = part.get_content_type()
                        content_disposition = str(part.get("Content-Disposition"))
                        try:
                            # get the email body
                            body = part.get_payload(decode=True).decode()
                        except:
                            pass
                        if content_type == "text/plain" and "attachment" not in content_disposition:
                            # get the date where the form is submitted
                            keyword = "Time Finished:"
                            position = body.find(keyword)
                            start_index = position + len(keyword) + 4
                            end_index = start_index + 2
                            date_list = ["", "", ""]
                            for j in range(0, 3):
                                date_list[j] = body[start_index:end_index]
                                start_index = end_index + 1;
                                end_index = start_index + 2;
                            date_str = date_list[1] + "/" + date_list[2] + "/" + date_list[0] + " - " + "lh"
                            # print(date_str)

                            # get the name of the trainee
                            if (real_type <= 1):
                                keyword = "Trainee/Intern Name"
                                second_key = "Host Name"
                            else:
                                keyword = "First Name (s)"
                                second_key = "Last Name (s)"
                            position = body.find(keyword)
                            start_index = position + len(keyword) + 3
                            end_index = body.find(' ', start_index)
                            first_name = body[start_index:end_index]
                            position = body.find(second_key)
                            start_index = position + len(second_key) + 3
                            end_index = body.find('\n', start_index)
                            name2 = body[start_index:end_index]
                            if real_type>1:
                                first_name = first_name[:(len(first_name)-2)]
                                name2 = name2[:(len(name2) - 1)]

                            # print(first_name, "and", name2)
                            list_of_trainee.append(Trainee(real_type, first_name, name2, date_str))

    imap.close()
    imap.logout()
    return list_of_trainee


def main():
    list = read_email(20)
    for i in range(0,len(list)):
        print(str(i), ". ", list[i].firstname)
        if list[i].type <= 1:
            print("\twith host", list[i].name2)
        else:
            print("\twith last name", list[i].name2)
        print("\ton date", list[i].date)

if __name__ == "__main__":
  main()
