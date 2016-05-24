#####
# Written by Gaddy
#####

#!/usr/bin/env python

import sys
import imaplib
import getpass

def process_mailbox(mail):
    connect, data = mail.search(None, "ALL")
    if connect != 'OK':
        print "No Emails found!"
        return

    for email in data[0].split():
        connect, data = mail.fetch(email, '(RFC822)')
        if connect != 'OK':
            print "ERROR getting message", email
            return
        print "Writing message ", email
        f = open('%s/%s.eml' %(<Output Directory>, email), 'wb')
        f.write(data[0][1])
        f.close()

def main():
    mail = imaplib.IMAP4_SSL('imap.gmail.com')
    mail.login(<EmailID>, getpass.getpass())
    connect, data = mail.select(<FolderName>)
    if connect == 'OK':
        print "Processing Folder: ", <FolderName>
        process_mailbox(mail)
        mail.close()
    else:
        print "ERROR: Unable to open Mailbox ", connect
    mail.logout()

if __name__ == "__main__":
    main()
