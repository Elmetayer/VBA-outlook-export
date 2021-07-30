# VBA-outlook-export
VBA macro executed in Outlook for exporting email details in csv format.
I've made this to elaborate statistics on my email sending and receiving in my professional environment through the years ...

Once executed in Outlook, this VBA code allows for exporting the details of all emails in a selected mailbox folder (including subfolders as the case may be).

Prerequisite: the macro is using RegExp, so you have to enable "Microsoft VBScript Regular Expressions" in the list of references (in the VBA editor, go to the menu Tools, then References, then select "Microsoft VBScript Regular Expressions x.x" where x.x is the current version on your system, e.g. 5.5)

When launched, the VBA macro asks to select any folder of your mailbox.
Once the folder is selected, the details of all emails in the selected folder (including its subfolders as the case may be) are exported in two csv files:
- "Export_Details_full", with all details for each recipient (several records for each email, depending on the number of recipients)
- "Export_Details_flat", without details for each recipient (one record for each email), that doesn’t contain the recipient information
A variable “export_allDetails”, when set to false, allows exporting data only in the Export_Details_flat csv file, this can be useful when exporting a lot of emails.

The CSV file names and folders are specified in the "fileNameDetails_all" and "fileNameDetails_from" variable: you will have to change this variable in order to successfully export your data.
In the "ExportAllEmail_Details" csv file, one csv record (i.e. line) is created for each recipient of a given email, so this can generate several csv records for a given email. Certain email attributes (e.g. sender, subject, etc.) are therefore repeated for each generated record. For each csv line, following information is exported in the "ExportAllEmail_Details" csv file:
- FROM : clean sender's name, ready for post-processing
- FROM_ADDRESS : sender's email address
- FROM_DOMAIN : sender's email domain
- FROM_NAME : sender's raw name
- TO : clean recipient's name, ready for post-processing /only in ExportAllEmail_Details_full file
- TO_ADDRESS : recipient's email address /only in ExportAllEmail_Details_full file
- TO_DOMAIN : recipient's email domain /in ExportAllEmail_Details_flat file, this field is set to "MULTI" in case the domains are different
- TO_NAME : recipient's raw name /only in ExportAllEmail_Details_full file
- TYPE : type of reception (TO or CC) /only in ExportAllEmail_Details_full file
- RECIPIENT_NUMBER : number of recipients
- RECIPIENT_NUMBER_TO : number of recipients (TO)
- RECIPIENT_NUMBER_CC : number of recipients (CC)
- DATE : date when email was sent
- HOUR : hour when email was sent
- DAY : day when email was sent
- WEEKDAY : weekday when email was sent
- WEEK : week number when email was sent
- YEAR : year when email was sent
- MONTH : month when email was sent
- SUBJECT : email's subject
- SUBJECT_WORDS : number of words in email's subject
- CONVERSATION : email’s converstation, i.e. subject without e.g. "RE:" or "FW:"
- BODY_WORDS : number of words in email's body
- URL_NUMBER : number of URLs in email's body
- EMAIL_NUMBER : number of email addresses in email's body
- ATTACHMENT_NUMBER : email's number of attachments
- ATTACHMENT_SIZE : email's total size of attachments
- EMAIL_ITEM_KEY : key representing the email

In addition, all email bodies (text) are aggregated in a “corpus” txt file. This file can be divided into multiple files in case the whole file is too big (see "maxSizeCorpus" parameter).

The exported files are organized as follows:
- a dedicated export folder named "ExportOutlook\yyyymmdd_hhmmss" is created for each export launch (yyyymmdd_hhmmss being the timestamp of extraction generation)
- the csv files and subfolders are put at the root of the above export folder
- each email body is exported in a txt file in a "Messages" subfolder, the name of each body txt file is the same as EMAIL_ITEM_KEY information (see hereafter)
- all corpus txt file(s) are exported in a “Corpus” subfolder

"Clean up text" steps are included in the code part where email bodies are exported, in order to remove specific items (recurrent emails signature, premises’ address, etc.) ... but this doesn’t work very well though, especially for signatures.
