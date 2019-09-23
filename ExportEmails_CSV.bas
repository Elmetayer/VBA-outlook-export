Attribute VB_Name = "ExportEmails_CSV"
Sub ExportToCSV()

'On Error GoTo ErrHandler
    Dim fileName As String
    Dim header As String
    Dim separator As String
    Dim counter As Long
    Dim intColumnCounter As Integer
    Dim msg As Outlook.MailItem
    Dim nms As Outlook.NameSpace
    Dim fld As Outlook.MAPIFolder
    Dim subFld As Outlook.MAPIFolder
    Dim itm As Object
    Dim OutputFileNum As Long
    Dim objStream
                
    'Open CSV file
    OutputFileNum = FreeFile
    fileName = "D:\Users\lemetayerbru\Downloads\ExportOutlook.csv"
    Open fileName For Output As #OutputFileNum
    
    'Definition of CSV separator
    separator = ";"
    
    'Definition of CSV file header
    'FROM : clean sender's name, ready for post-processing
    'FROM_ADDRESS : sender's email address
    'FROM_DOMAIN : sender's email domain
    'FROM_NAME : sender's raw name
    'TO : clean recipient's name, ready for post-processing
    'TO_ADDRESS : recipient's email address
    'TO_DOMAIN : sender's email domain
    'TO_NAME : recipient's raw name
    'TYPE : type of reception (TO or CC)
    'RECIPIENT_NUMBER
    'DATE : date when email was sent
    'HOUR : hour when email was sent
    'DAY : day when email was sent
    'MONTH : month when email was sent
    'SUBJECT : email's subject
    'SUBJECT_WORDS : number of words in email's subject
    'BODY_WORDS : number of words in email's body
    'URL_NUMBER : number of URLs in email's body
    'EMAIL_NUMBER : number of email addresses in email's body
    'ATTACHMENT_NUMBER : email's number of attachments
    'ATTACHMENT_SIZE : email's total size of attachments
    header = "FROM" & separator & "FROM_ADDRESS" & separator & "FROM_DOMAIN" & separator & "FROM_NAME" & separator & "TO" & separator & "TO_ADDRESS" & separator & _
    "TO_DOMAIN" & separator & "TO_NAME" & separator & "TYPE" & separator & "RECIPIENT_NUMBER" & separator & "DATE" & separator & "HOUR" & separator & "DAY" & separator & _
    "MONTH" & separator & "SUBJECT" & separator & "SUBJECT_WORDS" & separator & "BODY_WORDS" & separator & "URL_NUMBER" & separator & "EMAIL_NUMBER" & separator & _
    "ATTACHMENT_NUMBER" & separator & "ATTACHMENT_SIZE"
    
    'Print csv header
    Print #OutputFileNum, header;
    
    'Select Outlook folder to export
    Set nms = Application.GetNamespace("MAPI")
    Set fld = nms.PickFolder
    
    'Initialize counter of processed emails
    counter = 0
    
    'Handle potential errors with Select Folder dialog box
    If fld Is Nothing Then
        'Close CSV file
        Close OutputFileNum
        MsgBox "There are no mail messages to export", vbOKOnly, "Error"
        Exit Sub
    ElseIf fld.DefaultItemType <> olMailItem Then
        'Close CSV file
        Close OutputFileNum
        MsgBox "There are no mail messages to export", vbOKOnly, "Error"
        Exit Sub
    ElseIf fld.Items.Count = 0 And fld.Folders.Count = 0 Then
        'Close CSV file
        Close OutputFileNum
        MsgBox "There are no mail messages to export", vbOKOnly, "Error"
        Exit Sub
    ElseIf fld.Items.Count = 0 And fld.Folders.Count > 0 Then
        For Each subFld In fld.Folders
            For Each itm In subFld.Items
                'Write email field subfolder items in XML file
                If TypeOf itm Is MailItem Then
                    Set msg = itm
                    Print #OutputFileNum, parseEmailItem(msg);
                    counter = counter + 1
                End If
            Next itm
        Next subFld
        'Close CSV file
        Close OutputFileNum
        MsgBox "Done! " & counter & " element(s) processed", vbOKOnly
        Exit Sub
    ElseIf fld.Items.Count > 0 And fld.Folders.Count > 0 Then
        For Each itm In fld.Items
            'Write email field folder items in XML file
            If TypeOf itm Is MailItem Then
                Set msg = itm
                Print #OutputFileNum, parseEmailItem(msg);
                counter = counter + 1
            End If
        Next itm
        For Each subFld In fld.Folders
            For Each itm In subFld.Items
                'Write email field subfolder items in XML file
                If TypeOf itm Is MailItem Then
                    Set msg = itm
                    Print #OutputFileNum, parseEmailItem(msg);
                    counter = counter + 1
                End If
            Next itm
        Next subFld
        'Close CSV file
        Close OutputFileNum
        MsgBox "Done! " & counter & " element(s) processed", vbOKOnly
        Exit Sub
     ElseIf fld.Items.Count > 0 And fld.Folders.Count = 0 Then
        For Each itm In fld.Items
            'Write email field folder items in XML file
            If TypeOf itm Is MailItem Then
                Set msg = itm
                Print #OutputFileNum, parseEmailItem(msg);
                counter = counter + 1
            End If
        Next itm
        'Close CSV file
        Close OutputFileNum
        MsgBox "Done! " & counter & " element(s) processed", vbOKOnly
        Exit Sub
    End If
    
    Set msg = Nothing
    Set nms = Nothing
    Set fld = Nothing
    Set itm = Nothing

Exit Sub

ErrHandler:
    If Err.Number = 1004 Then
        MsgBox strSheet & " doesn't exist", vbOKOnly, "Error"
    Else
        MsgBox Err.Number & "; Description: " & Err.Description, vbOKOnly, "Error"
    End If

    Set msg = Nothing
    Set nms = Nothing
    Set fld = Nothing
    Set itm = Nothing

End Sub

Function parseEmailItem(msg As MailItem) As String
'Export a given email as CSV file records
'One record is created for each recipient, so this function can generate several CSV records
'Certain email attributes (e.g. sender, subject, etc.) are therefore repeated for each generated record

Dim i As Integer
Dim arr() As String
Dim urls As collection
Dim emailaddresses As collection
Dim separator As String
Dim cleanbody As String
Dim cleanupbody As String
Dim recipient As recipient
Dim recipientCount As Integer
Dim recipientType As String
Dim MailItem_FROM As String
Dim MailItem_FROM_ADDRESS As String
Dim MailItem_FROM_DOMAIN As String
Dim MailItem_FROM_NAME As String
Dim MailItem_TO As String
Dim MailItem_TO_ADDRESS As String
Dim MailItem_TO_DOMAIN As String
Dim MailItem_TO_NAME As String
Dim MailItem_TYPE As String
Dim MailItem_RECIPIENT_NUMBER As String
Dim MailItem_DATE As String
Dim MailItem_DAY As Integer
Dim MailItem_HOUR As Integer
Dim MailItem_MONTH As Integer
Dim MailItem_RTE_FLOW As String
Dim MailItem_SUBJECT As String
Dim MailItem_SUBJECT_WORDS As String
Dim MailItem_BODY_WORDS As String
Dim MailItem_URL_NUMBER As String
Dim MailItem_ATTACHMENT_NUMBER As String
Dim MailItem_ATTACHMENT_SIZE As String
Dim MailItem_EMAIL_NUMBER As String
Dim carriage_return As String
Dim mail_content As String
Dim mail_collection As String


'Definition of CSV file carriage return
carriage_return = Chr(13) & Chr(10)

'Definition of CSV file separator
separator = ";"

'Sender information
If msg.Sender Is Nothing Then
    'This case can happen sometimes in Outlook ...
    MailItem_FROM_DOMAIN = "EMPTY_SENDER"
    MailItem_FROM_NAME = "EMPTY_SENDER"
    MailItem_FROM_ADDRESS = "EMPTY_SENDER"
    MailItem_FROM = "EMPTY_SENDER"
Else
    'Domain information
    MailItem_FROM_DOMAIN = getDomainFromEmailAddress(cleanupEmailAddress(msg.SenderEmailAddress))
    'Raw name information
    'For email sent from outside, this information isn't always ready to be processed
    MailItem_FROM_NAME = cleanupEmailName(msg.SenderName)
    'Email address information
    MailItem_FROM_ADDRESS = cleanupEmailAddress(msg.SenderEmailAddress)
    'The "FROM" information is meant to be the one used in post processing
    'So the goal is to have something as close as possible to a "name" information
    If MailItem_FROM_DOMAIN = "RTE" Then
        'For emails sent from within your organization, the email name information is clean and can be used as such
        MailItem_FROM = MailItem_FROM_NAME
    Else
        'Emails sent from outside don't necessary have a "clean" email name
        'So, for better post processing, it's safer to use directly the "user" information from the email address
        'E.g. "surname.name" for an email address "surname.name@domain.com"
        MailItem_FROM = getUserFromEmailAddress(MailItem_FROM_ADDRESS)
    End If
End If

'Subject information
'Subject
MailItem_SUBJECT = cleanupText(msg.Subject)
'Subject word count
arr = VBA.Split(msg.Subject, " ")
MailItem_SUBJECT_WORDS = CStr(UBound(arr) - LBound(arr) + 1)

'Date information
MailItem_DATE = msg.SentOn
MailItem_DAY = Weekday(MailItem_DATE, vbMonday)
MailItem_HOUR = Hour(MailItem_DATE)
MailItem_MONTH = Month(MailItem_DATE)

'Recipient number information
MailItem_RECIPIENT_NUMBER = msg.Recipients.Count

'Body information, not taking into account the email history
'Only the "meaningful" body is used
'Body word count
cleanupbody = cleanupText(msg.body)
cleanbody = deleteEmailHistory(cleanupbody)
cleanbody_URLs_emails = deleteEmailEmailAddresses(deleteEmailURLs(cleanbody))
arr = VBA.Split(cleanbody_URLs_emails, " ")
MailItem_BODY_WORDS = CStr(UBound(arr) - LBound(arr) + 1)
'Body URL number count
Set urls = getEmailURLs(cleanbody)
MailItem_URL_NUMBER = urls.Count
'Body email number count
Set emailaddresses = getEmailEmailAddresses(cleanbody)
MailItem_EMAIL_NUMBER = emailaddresses.Count

'Attachments information
'Number of attachment
MailItem_ATTACHMENT_NUMBER = msg.Attachments.Count
'Total size of attachments
MailItem_ATTACHMENT_SIZE = 0
'The calculation is only made if there is at least one attachment
If msg.Attachments.Count > 0 Then
    For i = 0 To msg.Attachments.Count - 1
        MailItem_ATTACHMENT_SIZE = MailItem_ATTACHMENT_SIZE + msg.Attachments.Item(i + 1).Size
    Next i
End If

'Initialization of the record set to export
mail_collection = ""

'Creation of one record per recipient
For Each recipient In msg.Recipients
    'Initialization of the record
    mail_content = ""
    
    'Recipient information, same as for sender information
    'Domain information
    MailItem_TO_DOMAIN = getDomainFromEmailAddress(cleanupEmailAddress(recipient.Address))
    'Raw name information
    'For email sent to outside, this information isn't always ready to be processed
    MailItem_TO_NAME = cleanupEmailName(recipient.name)
    'Email address information
    MailItem_TO_ADDRESS = cleanupEmailAddress(recipient.Address)
    'The "TO" information is meant to be the one used in post processing
    'So the goal is to have something as close as possible to a "name" information
    If MailItem_TO_DOMAIN = "RTE" Then
        'For emails sent from within your organization, the email name information is clean and can be used as such
        MailItem_TO = MailItem_TO_NAME
    Else
        'Emails sent from outside don't necessary have a "clean" email name
        'So, for better post processing, it's safer to use directly the "user" information from the email address
        'E.g. "surname.name" for an email address "surname.name@domain.com"
        MailItem_TO = getUserFromEmailAddress(MailItem_TO_ADDRESS)
    End If
    
    'Reception type information, TO or CC
    If recipient.Type = 1 Then
        MailItem_TYPE = "TO"
    Else
        MailItem_TYPE = "CC"
    End If
    
    'Full record content
    mail_content = MailItem_FROM & separator & MailItem_FROM_ADDRESS & separator & MailItem_FROM_DOMAIN & separator & MailItem_FROM_NAME & separator & MailItem_TO & separator & _
    MailItem_TO_ADDRESS & separator & MailItem_TO_DOMAIN & separator & MailItem_TO_NAME & separator & MailItem_TYPE & separator & MailItem_RECIPIENT_NUMBER & separator & _
    MailItem_DATE & separator & MailItem_HOUR & separator & MailItem_DAY & separator & MailItem_MONTH & separator & MailItem_SUBJECT & separator & MailItem_SUBJECT_WORDS & separator & _
    MailItem_BODY_WORDS & separator & MailItem_URL_NUMBER & separator & MailItem_EMAIL_NUMBER & separator & MailItem_ATTACHMENT_NUMBER & separator & MailItem_ATTACHMENT_SIZE
    
    'Addition of the record content with a carriage return
    mail_collection = mail_collection & carriage_return & mail_content
    
Next recipient

parseEmailItem = mail_collection

End Function

Function deleteEmailHistory(text As String) As String
'Remove email history from a given email, based on a text representing the body

Dim replacementString As String
Dim temp As String
Dim regex As New VBScript_RegExp_55.RegExp

'Initialization of regex, especially with the pattern to use
'The pattern is set to correctly identify something that recalls the previous email sent in the same discussion thread
regex.Pattern = "(de|von|from)( {0,1}:).*(envoyé|gesendet|sent)( {0,1}:).*"
regex.Global = False
regex.MultiLine = True
regex.IgnoreCase = True

'Replacement string of regex
'It is set to keep a trace of deleted email history
replacementString = " [DELETED EMAIL HISTORY] "

'Execution of regex
'If there is a match, replace email history with the predefined replacement string
If regex.Test(text) Then
    temp = regex.Replace(text, replacementString)
Else
    temp = text
End If

Set regex = Nothing

deleteEmailHistory = temp

End Function

Function deleteEmailURLs(text As String) As String
'Remove URLs from a given email, based on a text representing the body

Dim replacementString As String
Dim temp As String
Dim regex As New VBScript_RegExp_55.RegExp

'Initialization of regex, especially with the pattern to use
'The pattern is set to correctly identify an URL
regex.Pattern = "(http[s]?|ftp):[/]+(\S+)"
regex.Global = True
regex.MultiLine = True
regex.IgnoreCase = True

'Replacement string of regex
'It is set to keep a trace of deleted URL
replacementString = " [DELETED URL] "

'Execution of regex
'If there is a match, replace URL with the predefined replacement string
If regex.Test(text) Then
    temp = regex.Replace(text, replacementString)
Else
    temp = text
End If

Set regex = Nothing

deleteEmailURLs = temp

End Function

Function deleteEmailEmailAddresses(text As String) As String
'Remove email addresses from a given text

Dim replacementString As String
Dim temp As String
Dim regex As New VBScript_RegExp_55.RegExp

'Initialization of regex, especially with the pattern to use
'The pattern is set to correctly identify an email address
regex.Pattern = "(mailto:)?(\S+)@(\S+)"
regex.Global = True
regex.MultiLine = True
regex.IgnoreCase = True

'Replacement string of regex
'It is set to keep a trace of deleted email address
replacementString = " [DELETED EMAIL ADDRESS] "

'Execution of regex
'If there is a match, replace email address with the predefined replacement string
If regex.Test(text) Then
    temp = regex.Replace(text, replacementString)
Else
    temp = text
End If

Set regex = Nothing

deleteEmailEmailAddresses = temp

End Function

Function getEmailURLs(text As String) As collection
'Return a collection with all URLs that are present in a given text

Dim temp As New collection
Dim regex As New VBScript_RegExp_55.RegExp

'Initialization of regex, especially with the pattern to use
'The pattern is set to correctly identify an URL
regex.Pattern = "(http[s]?|ftp):[/]+(\S+)"
regex.Global = True
regex.MultiLine = True
regex.IgnoreCase = True

'Execution of regex
'If there is a match, add URL to the collection
Set matches = regex.Execute(text)
For Each match In matches
    temp.Add (match.value)
Next match

Set regex = Nothing

Set getEmailURLs = temp
Set temp = Nothing

End Function

Function getEmailEmailAddresses(text As String) As collection
'Return a collection with all email addresses that are present in a given text

Dim temp As New collection
Dim regex As New VBScript_RegExp_55.RegExp

'Initialization of regex, especially with the pattern to use
'The pattern is set to correctly identify an email address
regex.Pattern = "(mailto:)?(\S+)@(\S+)"
regex.Global = True
regex.MultiLine = True
regex.IgnoreCase = True

'Execution of regex
'If there is a match, add email address to the collection
Set matches = regex.Execute(text)
For Each match In matches
    temp.Add (match.value)
Next match

Set regex = Nothing

Set getEmailEmailAddresses = temp
Set temp = Nothing

End Function

Function cleanupEmailAddress(text As String) As String
'This is essentially meant to remove "/CN=RECIPIENTS/CN= ...." from the email addresses that stem from your organization
'For emails sent from outside, this function returns normally the email as such

Dim replacementString As String
Dim temp As String
Dim regex As New VBScript_RegExp_55.RegExp

'Initialization of regex, especially with the pattern to use
'The pattern corresponds to the "messy" part of email addresses stemming from your organization
regex.Pattern = ".*(/CN=RECIPIENTS/CN=)(\w{32}-)?"
regex.Global = True
regex.MultiLine = True
regex.IgnoreCase = True

'Replacement string for regex
'As we only want to keep the "clean part" of the email address, we want to delete the rest
replacementString = ""
If regex.Test(text) Then
    temp = regex.Replace(text, replacementString)
Else
    temp = text
End If

Set regex = Nothing

'In some cases, it is needed to remove the '' that are put arround external email addresses
'e.g. 'surname.name@domain.com', to be changed in surname.name@domain.com
temp = Replace(temp, "'", "")
temp = Replace(temp, "’", "")

cleanupEmailAddress = temp

End Function

Function getDomainFromEmailAddress(text As String) As String
'Get domain from a given email address
'This can be a bit tricky in case you process an email address stemming from your organization

Dim replacementString As String
Dim temp As String
Dim companyDomain As String
Dim regex As New VBScript_RegExp_55.RegExp

'Initialization of your company's domain
companyDomain = "RTE"

'Initialization of regex, especially with the pattern to use
'The pattern is meant to identify "surname.name@" in an email address "surname.name@domain.com"
regex.Pattern = "(.+).(.+)@"
regex.Global = True
regex.MultiLine = True
regex.IgnoreCase = True

'Replacement string for regex
'As we only want to keep the domain, we want to delete the rest "surname.name@"
replacementString = ""

'Execution of the regex
If regex.Test(text) Then
    temp = regex.Replace(text, replacementString)
Else
    'The emails sent from within your organization often don't have the "@" address
    'They only have the username, e.g. SURNAME.NAME
    'This test allows to catch this case
    temp = companyDomain
End If

'The emails sent from within your organization sometimes have the "@" address
'But it is not necessary well formatted
'This test allows to catch this case and always assign the same company domain
If temp = "rte-france.com" Or temp = "RTE-France.com" Or temp = "rte-France.com" Then
    temp = companyDomain
End If

Set regex = Nothing

getDomainFromEmailAddress = temp

End Function
Function getUserFromEmailAddress(text As String) As String
'Get the "user" information from an email address
'E.g. "surname.name" for an email address "surname.name@domain.com"

Dim replacementString As String
Dim temp As String
Dim regex As New VBScript_RegExp_55.RegExp

'Initialization of regex, especially with the pattern to use
'The pattern is meant to identify "@domain.com" in an email address "surname.name@domain.com"
regex.Pattern = "'@(.*).(.*)"
regex.Global = True
regex.MultiLine = True
regex.IgnoreCase = True

'Replacement string for regex
'As we only want to keep the user name, we want to delete the rest "@domain.com"
replacementString = ""

If regex.Test(text) Then
    temp = regex.Replace(text, replacementString)
Else
    temp = text
End If

Set regex = Nothing

getUserFromEmailAddress = temp

End Function

Function cleanupText(text As String) As String
'Remove special character to grant a proper use of generated csv file
'This function is always used when processing a text information

Dim temp As String
temp = text

'List of special characters to be removed
temp = Replace(temp, Chr(13), " ")
temp = Replace(temp, Chr(10), " ")
temp = Replace(temp, Chr(13) + Chr(10), " ")
temp = Replace(temp, Chr(9), " ")
temp = Replace(temp, Chr(45), " ")
temp = Replace(temp, Chr(150), " ")
temp = Replace(temp, Chr(151), " ")
temp = Replace(temp, Chr(132), " ")
temp = Replace(temp, Chr(145), " ")
temp = Replace(temp, Chr(146), " ")
temp = Replace(temp, Chr(147), " ")
temp = Replace(temp, Chr(148), " ")
temp = Replace(temp, Chr(171), " ")
temp = Replace(temp, Chr(187), " ")

'List of other characters that aren't removed
'temp = Replace(temp, ",", " ")
'temp = Replace(temp, ";", " ")
'temp = Replace(temp, ".", " ")
'temp = Replace(temp, "<", " ")
'temp = Replace(temp, ">", " ")
'temp = Replace(temp, "/", " ")
'temp = Replace(temp, "\", " ")
'temp = Replace(temp, "'", " ")
'temp = Replace(temp, "’", " ")
'temp = Replace(temp, "?", " ")
'temp = Replace(temp, "!", " ")
'temp = Replace(temp, "=", " ")
'temp = Replace(temp, "-", " ")
'temp = Replace(temp, "_", " ")
'temp = Replace(temp, "–", " ")
'temp = Replace(temp, "ß", " ")
'temp = Replace(temp, "&", " ")

'Deletion of extra spaces (due to special characters' removal)
temp = removeExtraSpaces(temp)

cleanupText = temp

End Function

Function cleanupEmailName(name As String) As String
'Remove certain special characters from email name

Dim temp As String
temp = name

temp = Replace(temp, "'", "")
temp = Replace(temp, "’", "")

cleanupEmailName = temp

End Function
Function removeExtraSpaces(text As String)
'Deletion of extra spaces

Dim temp As String
temp = text

temp = Replace(temp, "  ", " ")
If temp = text Then
    removeExtraSpaces = temp
Else
    removeExtraSpaces = removeExtraSpaces(temp)
End If

End Function


