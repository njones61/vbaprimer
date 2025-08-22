Attribute VB_Name = "MailAppts"
Sub mail_appts()
'This sub looks at the appointments for the next week and generates an e-mail message
'with the list of appointments in an html-formatted e-mail message.
'The lines of code that need to be modified are marked in the comments.

'Get the list of appointments
Set myNamespace = ThisOutlookSession.GetNamespace("MAPI")
Set myCalendar = myNamespace.GetDefaultFolder(olFolderCalendar)
Set myItems = myCalendar.Items
myItems.IncludeRecurrences = True  'This ensures that recurring appts are fully included
myItems.Sort "[Start]"  'If you don't do this, they sometimes are listed in reverse order

'Create a new message
Set NewMail = ThisOutlookSession.CreateItem(olMailItem)

'***************** MODIFY THE NEXT LINE *********************
NewMail.Subject = "<insert title here> activities for the week of: " & _
       FormatDateTime(Now, vbShortDate) & _
   " - " & FormatDateTime(DateAdd("d", 7, Now), vbShortDate)

'***************** MODIFY THE NEXT LINE *********************
NewMail.Recipients.Add "<insert recipients here>"
NewMail.BodyFormat = olFormatHTML


'***************** MODIFY THE NEXT LINE *********************
NewMail.HTMLBody = NewMail.HTMLBody & "<body>" & "<p><font face=""Verdana"" size=""2"">" & _
   "Here <insert text here> activities for the coming week:</font></p>" & _
   "<hr>"


'Loop through the appointments adding the scout appointments from the next week
'to the body of the e-mail message.
Dim count As Integer
count = 0
For Each item In myItems

    'Check the date to see if it is within the next week
    thediff = DateDiff("d", Now, item.Start)
    
    If thediff >= 0 And thediff <= 8 Then
    
        
'***************** MODIFY THE NEXT LINE *********************
'IF YOU DON'T WANT TO CHECK ON A CATEGORY, REMOVE THIS IF STATEMENT
        'Check to see if it is a scouting related item
        If item.Categories = "<INSERT CATEGORY NAME>" Then
            count = count + 1
            
            If item.Location <> "" Then
                thelocation = "Location: " & item.Location & "<br>"
            Else
                thelocation = ""
            End If
            
            'Format the entry with the date/time/subject
            thedate = myFormatDate(item)
            thetime = myformattime(item)
            NewMail.HTMLBody = NewMail.HTMLBody & "<p><font face=""Verdana"" size=""2"">" & _
                "<b>" & item.Subject & "</b><br>" & _
                thelocation & _
                thedate & _
                thetime & _
                 "</font></p>"
            
            'Add the body notes if they exist
            If item.Body <> "" Then
                NewMail.HTMLBody = NewMail.HTMLBody & _
                "<p><font face=""Verdana"" size=""2""><i>" & item.Body & "</i></font></p>"
            Else
                NewMail.HTMLBody = NewMail.HTMLBody & "<br>"  'For spacing
            End If
            
            NewMail.HTMLBody = NewMail.HTMLBody & "<hr>"
        End If
        
    End If
    
Next item

'Add my name at the end

'***************** MODIFY THE NEXT LINE *********************
NewMail.HTMLBody = NewMail.HTMLBody & _
"<p><font face=""Verdana"" size = ""2"">Please contact me if you have any questions.</font></p>" & _
"<p><font face=""Verdana"" size=""2""><i>Warm regards,</i></font></p>" & _
"<p><font face=""Verdana"" size=""2"">Your Name<br>" & _
"<a href=""mailto:njones@byu.edu"">youraddress@yourdomain.com</a><br>" & _
"Your telephone number</font></p>" & _
"</body>"

NewMail.Display

End Sub

Function myFormatDate(item As Variant)
myFormatDate = "Date: " & FormatDateTime(item.Start, vbShortDate) & _
   " (" & WeekdayName(DatePart("w", item.Start)) & ")<br>"
End Function

Function myformattime(item As Variant)
Dim startTime As Variant
Dim endTime As Variant

'Format the times and remove the seconds part
startTime = Replace(FormatDateTime(item.Start, vbLongTime), ":00 ", " ")
endTime = Replace(FormatDateTime(item.End, vbLongTime), ":00 ", " ")
If Hour(item.Start) < 12 And Hour(item.End) < 12 Then
    'Remove the AM part from the start time.
    startTime = Replace(startTime, "AM", "")
ElseIf Hour(item.Start) >= 12 And Hour(item.End) >= 12 Then
    'Remove the PM part from the start time.
    startTime = Replace(startTime, "PM", "")
End If

myformattime = "Time: " & startTime & " - " & endTime & "<br>"

End Function
