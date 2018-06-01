import datetime
import email
import imaplib
import mailbox


EMAIL_ACCOUNT = "your_email_here"
PASSWORD = "password_here"

# connect to the gmail imap server
mail = imaplib.IMAP4_SSL('imap.gmail.com')

#login to your account(make sure you have allowed access to less secured apps in your Google account settings)
mail.login(EMAIL_ACCOUNT, PASSWORD)

mail.select('inbox') # You can select any mailbox here

# Put up your newsletter email address here
result, data = mail.uid('search', None, '(FROM "EMAIL ADDRESS HERE")')  # you could filter using the IMAP rules here (check http://www.example-code.com/csharp/imap-search-critera.asp)

i = len(data[0].split())

for x in range(i):
    latest_email_uid = data[0].split()[x]
    result, email_data = mail.uid('fetch', latest_email_uid, '(RFC822)') # fetching the mail, "`(RFC822)`" means "get the whole stuff", but you can ask for headers only, etc

    raw_email = email_data[0][1]
    raw_email_string = raw_email.decode('utf-8') # getting the mail content
    email_message = email.message_from_string(raw_email_string)

    # Header Details
    date_tuple = email.utils.parsedate_tz(email_message['Date'])
    if date_tuple:
        local_date = datetime.datetime.fromtimestamp(email.utils.mktime_tz(date_tuple))
        local_message_date = "%s" %(str(local_date.strftime("%a, %d %b %Y %H:%M:%S")))
    email_from = str(email.header.make_header(email.header.decode_header(email_message['From'])))
    email_to = str(email.header.make_header(email.header.decode_header(email_message['To'])))
    subject = str(email.header.make_header(email.header.decode_header(email_message['Subject'])))

# Body details
    for part in email_message.walk():
        if part.get_content_type() == "text/plain":
            body = part.get_payload(decode=True)
            file_name = str(x) + str(".") + subject[14:] + ".txt" 
            print("Writing the file " + file_name + " in your current directory. ")
            output_file = open(file_name, 'w')
            output_file.write("Body: \n\n%s" %(body.decode('utf-8',errors='ignore'))) # We ignore any unicode/decode errors we might encounter
            output_file.close()
        else:
            continue

