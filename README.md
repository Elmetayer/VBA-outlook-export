# VBA-outlook-export
VBA executed in Outlook for exporting email details in csv format

Once executed in Outlook, this VBA code allows for exporting all emails in a selected folder (including subfolders as the case may be).
For each email, following information is exported in a csv file:
FROM : clean sender's name, ready for post-processing
FROM_ADDRESS : sender's email address
FROM_DOMAIN : sender's email domain
FROM_NAME : sender's raw name
TO : clean recipient's name, ready for post-processing
TO_ADDRESS : recipient's email address
TO_DOMAIN : sender's email domain
TO_NAME : recipient's raw name
TYPE : type of reception (TO or CC)
RECIPIENT_NUMBER
DATE : date when email was sent
HOUR : hour when email was sent
DAY : day when email was sent
MONTH : month when email was sent
SUBJECT : email's subject
SUBJECT_WORDS : number of words in email's subject
BODY_WORDS : number of words in email's body
URL_NUMBER : number of URLs in email's body
EMAIL_NUMBER : number of email addresses in email's body
ATTACHMENT_NUMBER : email's number of attachments
ATTACHMENT_SIZE : email's total size of attachments

One csv record (i.e. line) is created for each recipient of a given email, so this function can generate several CSV records for a given email. Certain email attributes (e.g. sender, subject, etc.) are therefore repeated for each generated record.