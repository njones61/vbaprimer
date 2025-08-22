Attribute VB_Name = "email_utilities"
'This makes the code work with both 32- and 64-bit versions of Excel.
#If VBA7 Then
    Private Declare PtrSafe Function ShellExecute Lib _
                  "shell32.dll" Alias "ShellExecuteA" _
                  (ByVal hwnd As Long, _
                   ByVal lpOperation As String, _
                   ByVal lpFile As String, _
                   ByVal lpParameters As String, _
                   ByVal lpDirectory As String, _
                   ByVal nShowCmd As Long) As Long
#Else
    Private Declare Function ShellExecute Lib _
                  "shell32.dll" Alias "ShellExecuteA" _
                  (ByVal hwnd As Long, _
                   ByVal lpOperation As String, _
                   ByVal lpFile As String, _
                   ByVal lpParameters As String, _
                   ByVal lpDirectory As String, _
                   ByVal nShowCmd As Long) As Long
#End If
                
Private Const SW_SHOW = 1

Public Sub Navigate(ByVal NavTo As String)
  Dim hBrowse As Long
  hBrowse = ShellExecute(0&, "open", NavTo, "", "", SW_SHOW)
End Sub

Sub samp_browse()
'This illustrates how you can use the Navigate sub to launch a web page.
Navigate "http://www.vbcity.com"
End Sub

Sub send_mail(theaddress As String, _
              thesubject As String, _
               thebody As String)
'
'Send mail using the default mail client. Should work in every situation.
'


'This probably won't work for attachments or HTML formatted messages
'because it has a limit of about 1000 characters.  Will probably need
'to use outlook or more generic SMTP code.

'Navigate "mailto:njones@byu.edu?subject=" & _
         "There is something I want to tell you&body=" & _
         "This is the body of the message."
         
Navigate "mailto:" & theaddress & "?subject=" & _
         thesubject & "&body=" & _
         thebody

End Sub

Sub send_outlook_mail(tofield As String, _
                      ccfield As String, _
                      subjectfield As String, _
                      body As String, _
                      htmlformat As Boolean, _
                      attachment As String, _
                      displayfirst As Boolean)
                      
'This version should only be used when you have installed a copy of MS Outlook
'on your computer. It is more powerful than the above version and it lets
'you include attachments (provide the full path the file) and use HTML format if you like.

Set objOL = CreateObject("Outlook.Application")
Set objEmail = objOL.CreateItem(olMailItem)

With objEmail

    .To = tofield
    
    .CC = ccfield
    
    .Subject = subjectfield
    
    If htmlformat Then
        .HTMLBody = body
    Else
        .body = body
    End If
    
    .Importance = olImportanceNormal
    
    If attachment <> "" Then
        .attachments.Add attachment
    End If
    
    .Recipients.ResolveAll
    
    If displayfirst Then
        .Display
    Else
        .Send
    End If
    
End With

'MsgBox "Email Sent", vbOKOnly + vbInformation, "Send Mail Program"

End Sub

