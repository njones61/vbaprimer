Attribute VB_Name = "NewLetter"
Option Explicit


Sub SendLetterToContact()
Dim itmContact As Outlook.ContactItem
Dim selContacts As Selection
Dim objWord As Word.Application
Dim objLetter As Word.Document
Dim secNewArea As Word.Section

Set selContacts = Application.ActiveExplorer.Selection

If selContacts.Count > 0 Then
    Set objWord = New Word.Application
    
    For Each itmContact In selContacts
        
        Set objLetter = objWord.Documents.Add
        objLetter.Select

        With objLetter
            .Paragraphs.Add
            .Paragraphs.Add
            .Paragraphs.Add
            .Paragraphs.Add
            .Paragraphs.Add
            .Paragraphs.Add
        End With
        
        objWord.Selection.InsertAfter FormatDateTime(Now, vbLongDate)
        
         With objLetter
            .Paragraphs.Add
            .Paragraphs.Add
            .Paragraphs.Add
        End With
       
        objWord.Selection.InsertAfter itmContact.FullName
        objLetter.Paragraphs.Add
        
        If itmContact.CompanyName <> "" Then
            objWord.Selection.InsertAfter itmContact.CompanyName
            objLetter.Paragraphs.Add
        End If
        
        objWord.Selection.InsertAfter itmContact.BusinessAddress
        
        
        With objLetter
            .Paragraphs.Add
            .Paragraphs.Add
            .Paragraphs.Add
        End With
                
        objWord.Selection.InsertAfter "Dear " & itmContact.FirstName & ":"
        
        With objLetter
            .Paragraphs.Add
            .Paragraphs.Add
        End With
        
        objWord.Selection.InsertAfter "<Insert text of letter here>"
        
        With objLetter
            .Paragraphs.Add
            .Paragraphs.Add
            .Paragraphs.Add
        End With
        
        objWord.Selection.InsertAfter "Regards,"
        
        With objLetter
            .Paragraphs.Add
            .Paragraphs.Add
            .Paragraphs.Add
            .Paragraphs.Add
            .Paragraphs.Add
        End With
        
        objWord.Selection.InsertAfter Application.GetNamespace("MAPI").CurrentUser
    
    Next
    objWord.Visible = True
    
End If
End Sub
