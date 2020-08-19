import imaplib
import email
from email.header import decode_header
from read_from_excel import add_dates
from read_from_excel import Trainee
import webbrowser
import os

# Global variables
# List of possible types of emails
type_of_email = ["Host Trainer Mid-Point", "Host Trainer Final", "Trainee/Intern Initial",
                                 "Trainee/Intern Mid-Point", "Trainee/Intern Final"]


def read_email(number):
    # account credentials
    username = "abc@gmail.com"
    password = "somepassword"

    # create an IMAP4 class with SSL
    imap = imaplib.IMAP4_SSL("imap.gmail.com")
    # authenticate
    imap.login(username, password)
    # Open the mails under folder MAST of email
    status, messages = imap.select("MAST")
    # number of top emails to fetch
    N = number
    list_of_trainee = []
    # total number of emails
    messages = int(messages[0])

    # a list which elements contain a list of trainee name and a list of hostname
    combine_name = []
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
                # Check the type of the email
                # There are five types of emails, which can be classified
                # according to the subject of the email
                for index in range(0, len(type_of_email)):
                    if subject.find(type_of_email[index]) != -1:
                        real_type = index
                        break
                # This data entry email is has multipart
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
                            # Can be found with "Time Finished" headline
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

                            # get the name of the trainee and put them to a list
                            if (real_type <= 1):
                                keyword = "Trainee/Intern Name"
                            else:
                                keyword = "First Name (s)"
                            position = body.find(keyword)
                            start_index = position + len(keyword) + 3
                            end_index = body.find('\n', start_index)
                            trname = body[start_index:end_index]
                            trname.strip()
                            # If it is an evaluation from the trainee, have to add
                            # last name of the trainee from another entry
                            if (real_type > 1):
                                second_key = "Last Name (s)"
                                position = body.find(second_key)
                                start_index = position + len(second_key) + 3
                                end_index = body.find('\n', start_index)
                                lastname = body[start_index:end_index]
                                lastname.strip()
                                trname+=lastname
                            # Create trainee name list
                            tname = trname.title().split()
                            # print(tname)

                            # get the name of the host and put them to a list
                            if (real_type <= 1):
                                keyword = "Host Name"
                            else:
                                keyword = "Host Trainer"
                            position = body.find(keyword)
                            start_index = position + len(keyword) + 3
                            end_index = body.find('\n', start_index)
                            honame = body[start_index:end_index]
                            honame.strip()
                            # Create a list from the host name
                            hname = honame.title().split()
                            combine_name.append(Trainee(real_type,tname,hname,date_str))

    imap.close()
    imap.logout()
    return combine_name


def main():
    list = read_email(20)

    # Call the function add_dates with file_loc as
    # the location of file needed to be updated
    file_loc = "Progress Check-Ins 2020.xlsx"
    add_dates(list, file_loc)

    # Print out the report with trainee name, host name, type and date
    # also check if every trainee has been updated
    for index,trainee in enumerate(list):
        print("{}. Trainee Name: {}".format(index+1," ".join(trainee.tname)))
        print("\tEvaluation type: ", type_of_email[trainee.type])
        print("\tSubmitted date: {}".format(trainee.date))
        print("\tWith Host", " ".join(trainee.hname))
        if trainee.update == 0:
            print("\tNOT UPDATED")
        else:
            print("\tUpdated")

if __name__ == "__main__":
  main()
