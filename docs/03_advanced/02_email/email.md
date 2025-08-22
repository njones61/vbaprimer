# Sending E-Mail Using VBA in Excel

It is often useful to send e-mail using VB code. You can also send text messages using this technique. This page includes sample code and instructions.

## Adding the Email Utilities Module

First of all, you need to add some code to your project. Do the following:

1. Download the [email_utilities.bas](files/email_utilities.bas) file containing VB code for sending e-mail.
2. Go to your VB Editor window and right-click in the Project Explorer window on the left and select **Import File**.
3. Navigate to the **email_utilities.bas** file and open it. It will then be added as a new module.
4. Open the module and look at the functions and subs. You can now call these from anywhere in your code.

So to use the **send_mail** sub, you just need to call it and pass the arguments. Like this:

```vb
send_mail "myaddress@gmail.com", "TEST", "Hello world."
```

This will create a new message and launch your default e-mail client (Outlook, Thunderbird, etc). If you are working in the CAEDM network, there may not be an e-mail client set up on your account so the first time you run this you will probably need to go through a setup process. If you want to be able to send mail without having to hit the Send button on each message, you will need to use the Outlook version of the send_mail function.

## Sending Text Messages

Next you need to decide if you want to send e-mail or text messages. To send a text message you use the same code, but you use the phone number and carrier to formulate an e-mail address for the text message. First you have to determine the correct e-mail suffix from the following list:

| Company | E-mail tag |
|---------|------------|
| Sprint | messaging.sprintpcs.com |
| AT&T | mmode.com |
| Cingular | mobile.mycingular.com |
| Nextel | messaging.nextel.com |
| T-Mobile | tmomail.net |
| Verizon | vtext.com |

You can look up more complete lists on the internet. Then you take the phone number ("111-222-3333") and remove the dashes so that it is nothing but numbers ("1112223333") using the **Replace** string function. This creates the prefix. Then you combine the prefix and the suffix to generate the e-mail addresses as follows:

| Mobile # | Carrier | E-mail |
|----------|---------|--------|
| 801-111-1111 | Sprint | 8011111111@messaging.nextel.com |
| 801-222-2222 | Nextel | 8012222222@messaging.nextel.com |
| 801-333-3333 | Cingular | 8013333333@mobile.mycingular.com |
| 801-444-4444 | Nextel | 8014444444@messaging.nextel.com |
| 801-555-5555 | Nextel | 8015555555@messaging.nextel.com |
| 801-666-6666 | AT&T | 8016666666@mmode.com |
| 801-777-7777 | Verizon | 8017777777@vtext.com |
| 801-888-8888 | AT&T | 8018888888@mmode.com |
| 801-310-9291 | T-Mobile | 8013109291@tmomail.net |
| 801-209-5114 | AT&T | 8012095114@mmode.com |

Sometimes it is easiest to formulate the addresses just using Excel formulas. You can use a VLOOKUP to get the proper suffix from the carrier. Then you use your VB code to loop through the table and send your e-mail messages.

## Example VBA Code

Here's an example of how to send emails to multiple recipients using a loop:

```vb
Sub SendBulkEmails()
    Dim i As Integer
    Dim emailAddress As String
    Dim subject As String
    Dim message As String
    
    ' Set the subject and message
    subject = "Monthly Report"
    message = "Please find attached the monthly report for this month."
    
    ' Loop through a range of email addresses (assuming they're in column A starting at row 2)
    For i = 2 To Range("A" & Rows.Count).End(xlUp).Row
        emailAddress = Range("A" & i).Value
        
        ' Only send if there's an email address
        If emailAddress <> "" Then
            ' Send the email
            send_mail emailAddress, subject, message
        End If
    Next i
    
    MsgBox "Bulk email process completed!"
End Sub
```

## Text Message Example

Here's how to send text messages using the carrier email addresses:

```vb
Sub SendTextMessage()
    Dim phoneNumber As String
    Dim carrier As String
    Dim emailAddress As String
    Dim message As String
    
    ' Get phone number and carrier from user
    phoneNumber = InputBox("Enter phone number (e.g., 801-555-1234):", "Phone Number")
    carrier = InputBox("Enter carrier (Sprint, AT&T, Verizon, etc.):", "Carrier")
    message = InputBox("Enter your message:", "Message")
    
    ' Remove dashes from phone number
    phoneNumber = Replace(phoneNumber, "-", "")
    
    ' Create email address based on carrier
    Select Case LCase(carrier)
        Case "sprint"
            emailAddress = phoneNumber & "@messaging.sprintpcs.com"
        Case "at&t", "att"
            emailAddress = phoneNumber & "@mmode.com"
        Case "verizon"
            emailAddress = phoneNumber & "@vtext.com"
        Case "t-mobile", "tmobile"
            emailAddress = phoneNumber & "@tmomail.net"
        Case Else
            MsgBox "Carrier not recognized. Please check the spelling."
            Exit Sub
    End Select
    
    ' Send the text message
    send_mail emailAddress, "", message
    
    MsgBox "Text message sent to " & emailAddress
End Sub
```

## Advanced Email Features

The email utilities module provides several functions for different email scenarios:

### Sending with Attachments

```vb
' Send email with attachment
send_mail_with_attachment "recipient@example.com", "Subject", "Message", "C:\path\to\file.xlsx"
```

### Sending to Multiple Recipients

```vb
' Send to multiple recipients (comma-separated)
send_mail "user1@example.com,user2@example.com,user3@example.com", "Subject", "Message"
```

### HTML Email

```vb
' Send HTML formatted email
send_html_mail "recipient@example.com", "Subject", "<h1>Hello</h1><p>This is <b>HTML</b> formatted.</p>"
```

## Error Handling

It's important to add error handling when sending emails:

```vb
Sub SendEmailWithErrorHandling()
    On Error GoTo ErrorHandler
    
    Dim emailAddress As String
    emailAddress = "test@example.com"
    
    ' Attempt to send email
    send_mail emailAddress, "Test", "This is a test email"
    
    MsgBox "Email sent successfully!"
    Exit Sub
    
ErrorHandler:
    MsgBox "Error sending email: " & Err.Description
    Resume Next
End Sub
```

## Best Practices

1. **Always test with your own email first** - Make sure the system works before sending to others
2. **Use meaningful subject lines** - Helps recipients identify your emails
3. **Keep messages concise** - Text messages have character limits
4. **Handle errors gracefully** - Add error handling to prevent crashes
5. **Respect rate limits** - Don't send too many emails too quickly
6. **Verify email addresses** - Check that addresses are valid before sending

## Other Resources

For more information on sending email using VBA, see the following:

- [Ron de Bruin's Send Mail Page](http://www.rondebruin.nl/sendmail.htm) (great resource here - extensive set of sample code)
- [MSDN Outlook Sample Code](http://msdn.microsoft.com/en-us/library/office/ff458119%28v=office.11%29.aspx) (Outlook sample code)
- [MSDN Part 2](http://msdn.microsoft.com/en-us/library/office/ff519602%28v=office.11%29.aspx) (Part 2 from previous link)
- [MakeUseOf CDO Method](http://www.makeuseof.com/tag/send-emails-excel-vba/) (The CDO method - I haven't tried this, but it looks pretty slick)

## Troubleshooting

Common issues and solutions:

1. **Email client not configured** - Make sure you have a default email client set up
2. **Security warnings** - Some email clients may block automated emails
3. **Network restrictions** - Corporate networks may block certain email functionality
4. **Character encoding** - Special characters may not display correctly in some email clients

The email utilities module provides a simple and effective way to integrate email functionality into your Excel VBA applications, making it easy to automate communication tasks and send notifications automatically.
