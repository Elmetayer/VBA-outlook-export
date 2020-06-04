Attribute VB_Name = "ExportEmails_CSV"
Sub ExportEmailFullDetailsToCSV()

On Error GoTo ErrHandler
    Dim fileDirectory As String
    Dim fileDirectoryExport As String
    Dim fileDirectoryCorpus As String
    Dim fileNameDetails_all As String
    Dim fileNameDetails_from As String
    Dim fileNameCorpus As String
    Dim batchCorpusNumber As Integer
    Dim maxSizeCorpus As Long
    Dim currentSizeCorpus As Long
    Dim header_all As String
    Dim header_from As String
    Dim separator As String
    Dim counter As Long
    Dim intColumnCounter As Integer
    Dim msg As Outlook.MailItem
    Dim nms As Outlook.NameSpace
    Dim fld As Outlook.MAPIFolder
    Dim subFld As Outlook.MAPIFolder
    Dim itm As Object
    Dim OutputFileNumDetailsAll As Long
    Dim OutputFileNumDetailsFrom As Long
    Dim OutputFileNumCorpus As Long
    Dim parsedEmail_Output As collection
    Dim parsedEmail_Item_all As String
    Dim parsedEmail_Item_from As String
    Dim parsedEmail_Body As String
            
    'Set and create file directories
    'Parent file directory
    fileDirectory = "D:\Users\lemetayerbru\Downloads\ExportOutlook\" & Format(Now(), "yyyymmdd") & "_" & Format(Now(), "hhmmss") & "\"
    CreateDir (fileDirectory)
    'Subdirectory for individual email body
    fileDirectoryExport = fileDirectory & "\" & "Messages" & "\"
    CreateDir (fileDirectoryExport)
    'Subdirectory for all emails' corpus
    fileDirectoryCorpus = fileDirectory & "\" & "Corpus" & "\"
    CreateDir (fileDirectoryCorpus)
                
    'Open CSV file with all emails information
    OutputFileNumDetailsAll = FreeFile
    fileNameDetails_all = fileDirectory & "ExportAllEmail_Details.csv"
    Open fileNameDetails_all For Output As OutputFileNumDetailsAll
    
    'Open CSV file with all emails information for from only
    OutputFileNumDetailsFrom = FreeFile
    fileNameDetails_from = fileDirectory & "ExportEmail_DetailsFrom.csv"
    Open fileNameDetails_from For Output As OutputFileNumDetailsFrom

    'Set max size of txt file for corpus (if there are a lot of emails)
    maxSizeCorpus = 100000000
    'If the max size is exceeded, then another txt file is created, so there is a need of a batch counter
    batchCorpusNumber = 1
    currentSizeCorpus = 0
    'Open txt file with all corpus
    OutputFileNumCorpus = FreeFile
    fileNameCorpus = fileDirectoryCorpus & "ExportAllEmail_Words_" & batchCorpusNumber & ".txt"
    Open fileNameCorpus For Output As #OutputFileNumCorpus
        
    'Definition of CSV separator
    separator = ";"
    
    'Definition of CSV file header for all emails information
    'FROM : clean sender's name, ready for post-processing
    'FROM_ADDRESS : sender's email address
    'FROM_DOMAIN : sender's email domain
    'FROM_NAME : sender's raw name
    'TO : clean recipient's name, ready for post-processing
    'TO_ADDRESS : recipient's email address
    'TO_DOMAIN : sender's email domain
    'TO_NAME : recipient's raw name
    'TYPE : type of reception (TO or CC)
    'RECIPIENT_NUMBER : number of recipients
    'RECIPIENT_NUMBER_TO : number of recipients sent TO
    'RECIPIENT_NUMBER_CC : number of recipients sent CC
    'DATE : date when email was sent
    'HOUR : hour when email was sent
    'DAY : day when email was sent
    'WEEKDAY : weekday when email was sent
    'MONTH : month when email was sent
    'SUBJECT : email's subject, without "RE:" or "FW:"
    'CONVERSATION : contains element if the email is linked with a previous one, e.g. "RE:" or "FW:"
    'SUBJECT_WORDS : number of words in email's subject
    'BODY_WORDS : number of words in email's body
    'URL_NUMBER : number of URLs in email's body
    'EMAIL_NUMBER : number of email addresses in email's body
    'ATTACHMENT_NUMBER : email's number of attachments
    'ATTACHMENT_SIZE : email's total size of attachments (in bytes)
    'EMAIL_ITEM_KEY : key representing the email (see "MailItem_KEY" information in email parsing function)
    
    header_all = "FROM" & separator & "FROM_ADDRESS" & separator & "FROM_DOMAIN" & separator & "FROM_NAME" & separator & "TO" & separator & "TO_ADDRESS" & separator & "TO_DOMAIN" & separator & _
    "TO_NAME" & separator & "TYPE" & separator & "RECIPIENT_NUMBER" & separator & "RECIPIENT_NUMBER_TO" & separator & "RECIPIENT_NUMBER_CC" & separator & "DATE" & separator & "HOUR" & separator & "DAY" & separator & _
    "WEEKDAY" & separator & "MONTH" & separator & "SUBJECT" & separator & "CONVERSATION" & separator & "SUBJECT_WORDS" & separator & "BODY_WORDS" & separator & "URL_NUMBER" & separator & _
    "EMAIL_NUMBER" & separator & "ATTACHMENT_NUMBER" & separator & "ATTACHMENT_SIZE" & separator & "EMAIL_ITEM_KEY"
    
    'Print csv header for all emails information
    Print #OutputFileNumDetailsAll, header_all;

    'Definition of CSV file header for emails information for from only
    header_from = "FROM" & separator & "FROM_ADDRESS" & separator & "FROM_DOMAIN" & separator & "FROM_NAME" & separator & "RECIPIENT_NUMBER" & separator & "RECIPIENT_NUMBER_TO" & separator & _
    "RECIPIENT_NUMBER_CC" & separator & "DATE" & separator & "HOUR" & separator & "DAY" & separator & "WEEKDAY" & separator & "MONTH" & separator & "SUBJECT" & separator & _
    "CONVERSATION" & separator & "SUBJECT_WORDS" & separator & "BODY_WORDS" & separator & "URL_NUMBER" & separator & "EMAIL_NUMBER" & separator & "ATTACHMENT_NUMBER" & separator & _
    "ATTACHMENT_SIZE" & separator & "EMAIL_ITEM_KEY"
    
    'Print csv header for emails information for from only
    Print #OutputFileNumDetailsFrom, header_from;
    
    'Select Outlook folder to export
    Set nms = Application.GetNamespace("MAPI")
    Set fld = nms.PickFolder
    
    'Initialize counter of processed emails
    counter = 0
    
    'Handle potential errors with Select Folder dialog box
    If fld Is Nothing Then
        'Close files
        Close OutputFileNumDetailsAll
        Close OutputFileNumDetailsFrom
        Close OutputFileNumCorpus
        MsgBox "There are no mail messages to export", vbOKOnly, "Error"
        Exit Sub
    ElseIf fld.DefaultItemType <> olMailItem Then
        'Close files
        Close OutputFileNumDetailsAll
        Close OutputFileNumDetailsFrom
        Close OutputFileNumCorpus
        MsgBox "There are no mail messages to export", vbOKOnly, "Error"
        Exit Sub
    ElseIf fld.Items.Count = 0 And fld.Folders.Count = 0 Then
        'Close files
        Close OutputFileNumDetailsAll
        Close OutputFileNumDetailsFrom
        Close OutputFileNumCorpus
        MsgBox "There are no mail messages to export", vbOKOnly, "Error"
        Exit Sub
    ElseIf fld.Items.Count = 0 And fld.Folders.Count > 0 Then
        For Each subFld In fld.Folders
            For Each itm In subFld.Items
                'Write email field subfolder items in XML file
                If TypeOf itm Is MailItem Then
                    Set msg = itm
                    'Execution of parsing
                    Set parsedEmail_Output = parseEmailItem_allDetails(msg, fileDirectoryExport)
                    'Writing of all email details at the end of csv file
                    parsedEmail_Item_all = parsedEmail_Output.Item(1)
                    Print #OutputFileNumDetailsAll, parsedEmail_Item_all;
                    'Writing of email details (from only) at the end of csv file
                    parsedEmail_Item_from = parsedEmail_Output.Item(2)
                    Print #OutputFileNumDetailsFrom, parsedEmail_Item_from;
                    'Gathering of email's body
                    parsedEmail_Body = parsedEmail_Output.Item(3)
                    'Update of currentSizeCorpus to check whether it exceeds the max size
                    currentSizeCorpus = currentSizeCorpus + Len(parsedEmail_Body)
                    'If the size is exceeded, then create another corpus file
                    If currentSizeCorpus > maxSizeCorpus Then
                        Close OutputFileNumCorpus
                        batchCorpusNumber = batchCorpusNumber + 1
                        currentSizeCorpus = Len(parsedEmail_Body)
                        OutputFileNumCorpus = FreeFile
                        fileNameCorpus = fileDirectoryCorpus & "ExportAllEmail_Words_" & batchCorpusNumber & ".txt"
                        Open fileNameCorpus For Output As #OutputFileNumCorpus
                    End If
                    'Writing of email's body at the end of existing file
                    Print #OutputFileNumCorpus, parsedEmail_Body
                    counter = counter + 1
                End If
            Next itm
        Next subFld
        'Close files
        Close OutputFileNumDetailsAll
        Close OutputFileNumDetailsFrom
        Close OutputFileNumCorpus
        MsgBox "Done! " & counter & " element(s) processed" & Chr(13) & Chr(10) & _
        "Corpus files exported: " & batchCorpusNumber & Chr(13) & Chr(10) & _
        "Size of last corpus exported: " & currentSizeCorpus, vbOKOnly
        Exit Sub
    ElseIf fld.Items.Count > 0 And fld.Folders.Count > 0 Then
        For Each itm In fld.Items
            'Write email field folder items in XML file
            If TypeOf itm Is MailItem Then
                Set msg = itm
                'Execution of parsing
                Set parsedEmail_Output = parseEmailItem_allDetails(msg, fileDirectoryExport)
                'Writing of all email details at the end of csv file
                parsedEmail_Item_all = parsedEmail_Output.Item(1)
                Print #OutputFileNumDetailsAll, parsedEmail_Item_all;
                'Writing of email details (from only) at the end of csv file
                parsedEmail_Item_from = parsedEmail_Output.Item(2)
                Print #OutputFileNumDetailsFrom, parsedEmail_Item_from;
                'Gathering of email's body
                parsedEmail_Body = parsedEmail_Output.Item(3)
                'Update of currentSizeCorpus to check whether it exceeds the max size
                currentSizeCorpus = currentSizeCorpus + Len(parsedEmail_Body)
                'If the size is exceeded, then create another corpus file
                If currentSizeCorpus > maxSizeCorpus Then
                    Close OutputFileNumCorpus
                    batchCorpusNumber = batchCorpusNumber + 1
                    currentSizeCorpus = Len(parsedEmail_Body)
                    OutputFileNumCorpus = FreeFile
                    fileNameCorpus = fileDirectoryCorpus & "ExportAllEmail_Words_" & batchCorpusNumber & ".txt"
                    Open fileNameCorpus For Output As #OutputFileNumCorpus
                End If
                'Writing of email's body at the end of existing file
                Print #OutputFileNumCorpus, parsedEmail_Body
                counter = counter + 1
            End If
        Next itm
        For Each subFld In fld.Folders
            For Each itm In subFld.Items
                'Write email field subfolder items in XML file
                If TypeOf itm Is MailItem Then
                    Set msg = itm
                    'Execution of parsing
                    Set parsedEmail_Output = parseEmailItem_allDetails(msg, fileDirectoryExport)
                    'Writing of all email details at the end of csv file
                    parsedEmail_Item_all = parsedEmail_Output.Item(1)
                    Print #OutputFileNumDetailsAll, parsedEmail_Item_all;
                    'Writing of email details (from only) at the end of csv file
                    parsedEmail_Item_from = parsedEmail_Output.Item(2)
                    Print #OutputFileNumDetailsFrom, parsedEmail_Item_from;
                    'Gathering of email's body
                    parsedEmail_Body = parsedEmail_Output.Item(3)
                    'Update of currentSizeCorpus to check whether it exceeds the max size
                    currentSizeCorpus = currentSizeCorpus + Len(parsedEmail_Body)
                    'If the size is exceeded, then create another corpus file
                    If currentSizeCorpus > maxSizeCorpus Then
                        Close OutputFileNumCorpus
                        batchCorpusNumber = batchCorpusNumber + 1
                        currentSizeCorpus = Len(parsedEmail_Body)
                        OutputFileNumCorpus = FreeFile
                        fileNameCorpus = fileDirectoryCorpus & "ExportAllEmail_Words_" & batchCorpusNumber & ".txt"
                        Open fileNameCorpus For Output As #OutputFileNumCorpus
                    End If
                    'Writing of email's body at the end of existing file
                    Print #OutputFileNumCorpus, parsedEmail_Body
                    counter = counter + 1
                End If
            Next itm
        Next subFld
        'Close files
        Close OutputFileNumDetailsAll
        Close OutputFileNumDetailsFrom
        Close OutputFileNumCorpus
        MsgBox "Done! " & counter & " element(s) processed" & Chr(13) & Chr(10) & _
        "Corpus files exported: " & batchCorpusNumber & Chr(13) & Chr(10) & _
        "Size of last corpus exported: " & currentSizeCorpus, vbOKOnly
        Exit Sub
     ElseIf fld.Items.Count > 0 And fld.Folders.Count = 0 Then
        For Each itm In fld.Items
            'Write email field folder items in XML file
            If TypeOf itm Is MailItem Then
                Set msg = itm
                'Execution of parsing
                Set parsedEmail_Output = parseEmailItem_allDetails(msg, fileDirectoryExport)
                'Writing of all email details at the end of csv file
                parsedEmail_Item_all = parsedEmail_Output.Item(1)
                Print #OutputFileNumDetailsAll, parsedEmail_Item_all;
                'Writing of email details (from only) at the end of csv file
                parsedEmail_Item_from = parsedEmail_Output.Item(2)
                Print #OutputFileNumDetailsFrom, parsedEmail_Item_from; 'Gathering of email's body
                parsedEmail_Body = parsedEmail_Output.Item(3)
                'Update of currentSizeCorpus to check whether it exceeds the max size
                currentSizeCorpus = currentSizeCorpus + Len(parsedEmail_Body)
                'If the size is exceeded, then create another corpus file
                If currentSizeCorpus > maxSizeCorpus Then
                    Close OutputFileNumCorpus
                    batchCorpusNumber = batchCorpusNumber + 1
                    currentSizeCorpus = Len(parsedEmail_Body)
                    OutputFileNumCorpus = FreeFile
                    fileNameCorpus = fileDirectoryCorpus & "ExportAllEmail_Words_" & batchCorpusNumber & ".txt"
                    Open fileNameCorpus For Output As #OutputFileNumCorpus
                    End If
                'Writing of email's body at the end of existing file
                Print #OutputFileNumCorpus, parsedEmail_Body
                counter = counter + 1
            End If
        Next itm
        'Close files
        Close OutputFileNumDetailsAll
        Close OutputFileNumDetailsFrom
        Close OutputFileNumCorpus
        MsgBox "Done! " & counter & " element(s) processed" & Chr(13) & Chr(10) & _
        "Corpus files exported: " & batchCorpusNumber & Chr(13) & Chr(10) & _
        "Size of last corpus exported: " & currentSizeCorpus, vbOKOnly
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

Function parseEmailItem_allDetails(msg As MailItem, fileDirectory As String) As collection
'The function writes a cleaned up email body in a txt file in the directory specified by "fileDirectory" parameter
'The function also generates three outputs
'Return as first output CSV file records (parsing), one record for each recipient, so this can generate several CSV records
'(certain email attributes - e.g. sender, subject, etc. - are repeated for each generated record)
'Returns as second output one record for the email, not taking into account the recipients
'Returns as third output the cleaned up body of email (the one writed in the directory as well)

Dim output_1 As String
Dim output_2 As String
Dim output_3 As String
Dim output As collection
Dim i As Integer
Dim arr() As String
Dim deletedElements As collection
Dim separator As String
Dim URL_Pattern As String
Dim Email_Pattern As String
Dim EmailHistory_Pattern As String
Dim Signature_Pattern As String
Dim PhoneNumber_Pattern As String
Dim Number_Pattern As String
Dim SingleCharacter_Pattern As String
Dim body As String
Dim cleanbody As String
Dim final_cleanedupbody As String
Dim cleanedupbody As String
Dim cleanbody_URLs As String
Dim cleanbody_emails As String
Dim cleanbody_URLs_emails As String
Dim cleanbody_URLs_emails_numbers As String
Dim cleanbody_URLs_emails_numbers_Signature As String
Dim cleanbody_noHistory As String
Dim cleanbody_nonWords As String
Dim cleanbody_nonWords_Letters As String
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
Dim MailItem_RECIPIENT_NUMBER_TO As String
Dim MailItem_RECIPIENT_NUMBER_CC As String
Dim MailItem_DATE As String
Dim MailItem_DAY As Integer
Dim MailItem_WEEKDAY As Integer
Dim MailItem_HOUR As Integer
Dim MailItem_MINUTE As Integer
Dim MailItem_SECOND As Integer
Dim MailItem_MONTH As Integer
Dim MailItem_YEAR As Integer
Dim MailItem_SUBJECT As String
Dim MailItem_CONVERSATION As String
Dim MailItem_SUBJECT_WORDS As String
Dim MailItem_BODY_WORDS As String
Dim MailItem_URL_NUMBER As String
Dim MailItem_ATTACHMENT_NUMBER As String
Dim MailItem_ATTACHMENT_SIZE As String
Dim MailItem_EMAIL_NUMBER As String
Dim carriage_return As String
Dim mail_content_detailed As String
Dim mail_content_flat As String
Dim mail_collection As String
Dim MailItem_KEY As String
Dim MailItem_fileName As String
Dim MailItem_OutputFileNum As Long

'Definition of CSV file carriage return
carriage_return = Chr(13) & Chr(10)

'Definition of CSV file separator
separator = ";"

'Definition of patterns
URL_Pattern = "\<{0,1}(http[s]?|ftp|onenote):.+\>{0,1}"
Email_Pattern = "(mailto:)?(\S+)@(\S+)"
'The pattern is set to correctly identify something that recalls the previous email sent in the same discussion thread
'This pattern is made to be executed once all line returns have been removed
EmailHistory_Pattern = "\s(de|De|von|Von|from|From)(\s{0,1}:).*(envoy�|Envoy�|gesendet|Gesendet|sent|Sent)(\s{0,1}:).*"
Signature_Pattern = "(WINDOW|Window)\s7(C|c),{0,1}\s(place|Place|PLACE)\s(du|DU|Du)\s(D�me|Dome|DOME),{0,1}\s92073\s(Paris|PARIS)\s(La|la|LA)\s(D�fense|Defense|DEFENSE)\s(cedex|CEDEX|Cedex)"
Number_Pattern = "[0-9]+"
SingleCharacter_Pattern = "\s(\D\s)+"
'Phone number pattern is not used anymore, it's simpler to remove numbers
'PhoneNumber_Pattern = "[A-Z]{0,1}(\({0,1}\+{0,1}\d{2}\){0,1}(.{0,1}))((\(\d\)){0,1}|\d)\d.{0,1}(\d{2}(.{0,1})){4}"

'Parsing sender information
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

'Parsing Email's subject information
'Subject
MailItem_SUBJECT = cleanupText(msg.Subject)
'Conversation
MailItem_CONVERSATION = cleanupText(msg.ConversationTopic)
'Subject word count
arr = VBA.Split(msg.Subject, " ")
MailItem_SUBJECT_WORDS = CStr(UBound(arr) - LBound(arr) + 1)

'Parsing date information
MailItem_DATE = msg.SentOn
MailItem_DAY = Day(MailItem_DATE)
MailItem_WEEKDAY = Weekday(MailItem_DATE, vbMonday)
MailItem_MONTH = Month(MailItem_DATE)
MailItem_YEAR = Year(MailItem_DATE)
MailItem_HOUR = Hour(MailItem_DATE)
MailItem_MINUTE = Minute(MailItem_DATE)
MailItem_SECOND = Second(MailItem_DATE)

'Parsing recipient number information
MailItem_RECIPIENT_NUMBER = msg.Recipients.Count
MailItem_RECIPIENT_NUMBER_TO = 0
MailItem_RECIPIENT_NUMBER_CC = 0
For Each recipient In msg.Recipients
    'Reception type information, TO or CC
    If recipient.Type = 1 Then
        MailItem_RECIPIENT_NUMBER_TO = MailItem_RECIPIENT_NUMBER_TO + 1
    Else
        MailItem_RECIPIENT_NUMBER_CC = MailItem_RECIPIENT_NUMBER_CC + 1
    End If
Next recipient

'Parsing body information, not taking into account the email history
'Cleaning up the body to get only the "meaningful" body
body = msg.body

Set deletedElements = deleteEmailPattern(body, URL_Pattern, "")
'remove URLs
cleanbody_URLs = deletedElements.Item(1)
'Body email number count
MailItem_URL_NUMBER = deletedElements.Item(2)

'remove email addresses
Set deletedElements = deleteEmailPattern(cleanbody_URLs, Email_Pattern, "")
'Body email number count
MailItem_EMAIL_NUMBER = deletedElements.Item(2)

'intermediate clean up
cleanedupbody = cleanupSpecialCharacters(deletedElements.Item(1))

'remove numbers
Set deletedElements = deleteEmailPattern(cleanedupbody, Number_Pattern, "")

'remove email signature
Set deletedElements = deleteEmailPattern(deletedElements.Item(1), Signature_Pattern, "")

'remove history
Set deletedElements = deleteEmailPattern(deletedElements.Item(1), EmailHistory_Pattern, "")

'remove single characters
Set deletedElements = deleteEmailPattern(cleanupText(deletedElements.Item(1)), SingleCharacter_Pattern, " ")
final_cleanedupbody = deletedElements.Item(1)

'Body word count
'Create an array with each word
arr = VBA.Split(final_cleanedupbody, " ")
'Number of words in body is equal to the number of elements in the array
MailItem_BODY_WORDS = CStr(UBound(arr) - LBound(arr) + 1)

'Export body content in a separate txt file
'Create key used to name the exported txt file
MailItem_KEY = Format(MailItem_DATE, "yyyymmdd") & "_" & Format(MailItem_DATE, "hhmmss") & "_" & MailItem_FROM_NAME & "_" & MailItem_BODY_WORDS

'Open txt file to write email's body
MailItem_OutputFileNum = FreeFile
'Write email's content (cleanedupbody)
MailItem_fileName = fileDirectory & MailItem_KEY & ".txt"
Open MailItem_fileName For Output As #MailItem_OutputFileNum
Print #MailItem_OutputFileNum, final_cleanedupbody;

'Close txt file
Close MailItem_OutputFileNum

'Parsing attachments information
'Number of attachment
MailItem_ATTACHMENT_NUMBER = msg.Attachments.Count
'Total size of attachments
MailItem_ATTACHMENT_SIZE = 0
'The calculation is only made if there is at least one attachment
If msg.Attachments.Count > 0 Then
    For i = 0 To msg.Attachments.Count - 1
        'Size property returns a the size (in bytes) of the Outlook item
        MailItem_ATTACHMENT_SIZE = MailItem_ATTACHMENT_SIZE + msg.Attachments.Item(i + 1).Size
    Next i
End If

'Initialization of the unique record not taking from information
mail_content_flat = ""

'Initialization of the record set to export
mail_collection = ""

'Creation of one record per recipient
For Each recipient In msg.Recipients
    'Initialization of the detailed record
    mail_content_detailed = ""
    
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
    mail_content_detailed = MailItem_FROM & separator & MailItem_FROM_ADDRESS & separator & MailItem_FROM_DOMAIN & separator & MailItem_FROM_NAME & separator & MailItem_TO & separator & _
    MailItem_TO_ADDRESS & separator & MailItem_TO_DOMAIN & separator & MailItem_TO_NAME & separator & MailItem_TYPE & separator & MailItem_RECIPIENT_NUMBER & separator & _
    MailItem_RECIPIENT_NUMBER_TO & separator & MailItem_RECIPIENT_NUMBER_CC & separator & MailItem_DATE & separator & MailItem_HOUR & separator & MailItem_DAY & separator & _
    MailItem_WEEKDAY & separator & MailItem_MONTH & separator & MailItem_SUBJECT & separator & MailItem_CONVERSATION & separator & MailItem_SUBJECT_WORDS & separator & _
    MailItem_BODY_WORDS & separator & MailItem_URL_NUMBER & separator & MailItem_EMAIL_NUMBER & separator & MailItem_ATTACHMENT_NUMBER & separator & MailItem_ATTACHMENT_SIZE & _
    separator & MailItem_KEY
    
    'Addition of the record content with a carriage return
    mail_collection = mail_collection & carriage_return & mail_content_detailed
    
Next recipient

'Email record without taking into account from information
mail_content_flat = carriage_return & MailItem_FROM & separator & MailItem_FROM_ADDRESS & separator & MailItem_FROM_DOMAIN & separator & MailItem_FROM_NAME & separator & _
MailItem_RECIPIENT_NUMBER & separator & MailItem_RECIPIENT_NUMBER_TO & separator & MailItem_RECIPIENT_NUMBER_CC & separator & MailItem_DATE & separator & MailItem_HOUR & separator & _
MailItem_DAY & separator & MailItem_WEEKDAY & separator & MailItem_MONTH & separator & MailItem_SUBJECT & separator & MailItem_CONVERSATION & separator & MailItem_SUBJECT_WORDS & separator & _
MailItem_BODY_WORDS & separator & MailItem_URL_NUMBER & separator & MailItem_EMAIL_NUMBER & separator & MailItem_ATTACHMENT_NUMBER & separator & MailItem_ATTACHMENT_SIZE & separator & MailItem_KEY

'Creation of function's output
Set output = New collection

'First output is emails's parsed details
output_1 = mail_collection
output.Add (output_1)
'Second output is information not taking into account the recipients
output_2 = mail_content_flat
output.Add (output_2)
'Third output is email's body
'Note that we have to add a space in order to merge all bodies one after another in the final file
output_3 = final_cleanedupbody & " "
output.Add (output_3)
Set parseEmailItem_allDetails = output

End Function

Function deleteEmailPattern(text As String, pattern As String, replacementString As String) As collection
'Remove URLs from a given text
'Returns as first argument the cleaned text, without URLs
'Returns as second argument the number of deleted URLs

Dim output_1 As String
Dim output_2 As Long
Dim output As collection
Dim deletedEmailURLs As collection
Dim regex As New VBScript_RegExp_55.RegExp

Set output = New collection

'Initialization of regex, especially with the pattern to use
'The pattern is set to correctly identify an email address
regex.pattern = pattern
regex.Global = True
regex.MultiLine = True
regex.IgnoreCase = True

'Execution of regex
'If there is a match, add email address to the collection
Set deletedEmailURLs = New collection
Set matches = regex.Execute(text)
For Each match In matches
    deletedEmailURLs.Add (match.value)
Next match
output_2 = deletedEmailURLs.Count

'Execution of regex
'If there is a match, replace email address with the predefined replacement string
If regex.Test(text) Then
    output_1 = regex.Replace(text, replacementString)
Else
    output_1 = text
End If

Set regex = Nothing

output.Add (output_1)
output.Add (output_2)
Set deleteEmailPattern = output

End Function

Function cleanupEmailAddress(text As String) As String
'This is essentially meant to remove "/CN=RECIPIENTS/CN= ...." from the email addresses that stem from your organization
'For emails sent from outside, this function returns normally the email as such

Dim replacementString As String
Dim temp As String
Dim regex As New VBScript_RegExp_55.RegExp

'Initialization of regex, especially with the pattern to use
'The pattern corresponds to the "messy" part of email addresses stemming from your organization
regex.pattern = ".*(/CN=RECIPIENTS/CN=)(\w{32}-)?"
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
temp = Replace(temp, "�", "")

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
regex.pattern = "(.+).(.+)@"
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
regex.pattern = "'@(.*).(.*)"
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

Function cleanupSpecialCharacters(text As String) As String
'Remove special character to grant a proper use of generated csv file
'This function is always used when processing a text information

Dim temp As String
Dim myEmailSignature As String

temp = text

'my email's signature
myEmailSignature = "Bruno LEMETAYER PILOTE D'AFFAIRES DIES Direction Exploitation D�partement Exploitation Immeuble WINDOW 7C Place du D�me 92073 Paris La D�fense cedex"
temp = Replace(temp, myEmailSignature, "")

'List of special characters to be removed
temp = Replace(temp, Chr(13), " ")
temp = Replace(temp, Chr(10), " ")
temp = Replace(temp, Chr(13) + Chr(10), " ")
temp = Replace(temp, Chr(9), " ")

'List of other characters that aren't removed


'Deletion of extra spaces (due to special characters' removal)
temp = removeExtraSpaces(temp)

cleanupSpecialCharacters = temp

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
temp = Replace(temp, ",", " ")
temp = Replace(temp, ";", " ")
temp = Replace(temp, ":", " ")
temp = Replace(temp, ".", " ")
temp = Replace(temp, "<", " ")
temp = Replace(temp, ">", " ")
temp = Replace(temp, "/", " ")
temp = Replace(temp, "\", " ")
temp = Replace(temp, "?", " ")
temp = Replace(temp, "!", " ")
temp = Replace(temp, "=", " ")
temp = Replace(temp, "-", " ")
temp = Replace(temp, "_", " ")
temp = Replace(temp, "�", " ")
temp = Replace(temp, "�", " ")
temp = Replace(temp, "&", " ")
temp = Replace(temp, "*", " ")
temp = Replace(temp, "+", " ")
temp = Replace(temp, "�", " ")
temp = Replace(temp, "�", " ")
temp = Replace(temp, "(", " ")
temp = Replace(temp, ")", " ")
temp = Replace(temp, "@", " ")
temp = Replace(temp, "�", " ")

'List of other characters that aren't removed
'temp = Replace(temp, "'", " ")
'temp = Replace(temp, "�", " ")

'Deletion of extra spaces (due to special characters' removal)
temp = removeExtraSpaces(temp)

cleanupText = temp

End Function
Function cleanupEmailName(name As String) As String
'Remove certain special characters from email name

Dim temp As String
temp = name

temp = Replace(temp, "'", "")
temp = Replace(temp, "�", "")

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

Sub CreateDir(strPath As String)
    'strPath shall not include a filename after the final "\" or this code will create a folder with that name
    Dim elm As Variant
    Dim strCheckPath As String

    strCheckPath = ""
    For Each elm In Split(strPath, "\")
        strCheckPath = strCheckPath & elm & "\"
        If Len(Dir(strCheckPath, vbDirectory)) = 0 Then MkDir strCheckPath
    Next
End Sub

Sub CheckIfFileOpen(fileName As String, fileStream As Long)

'Call function to check if the file is open
If IsFileOpen(fileName) = True Then
    Close fileStream
End If

End Sub

Function IsFileOpen(fileName As String)

Dim fileNum As Integer
Dim errNum As Integer

'Allow all errors to happen
On Error Resume Next
fileNum = FreeFile()

'Try to open and close the file for input.
'Errors mean the file is already open
Open fileName For Input Lock Read As #fileNum
Close fileNum

'Get the error number
errNum = Err

'Do not allow errors to happen
On Error GoTo 0

'Check the Error Number
Select Case errNum

    'errNum = 0 means no errors, therefore file closed
    Case 0
    IsFileOpen = False
 
    'errNum = 70 means the file is already open
    Case 70
    IsFileOpen = True

    'Something else went wrong
    Case Else
    IsFileOpen = errNum

End Select

End Function
