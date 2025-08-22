Attribute VB_Name = "get_file"
Option Explicit

Function GetExportFileName(Filt As String, FilterIndex As String, defname As String, Prompt As String) As String
'
'Selects a file for export
'
'Use the following format to set up the filter
'    Filt = "Text Files (*.txt),*.txt," & _
'           "Lotus Files (*.prn),*.prn," & _
'           "Comma Separated Files (*.csv),*.csv," & _
'           "ASCII Files (*.asc),*.asc," & _
'           "XML Files (*.xml),*.xml," & _
'           "All Files (*.*),*.*"
'
' The Filter Index controls what filter is selected by default.  It starts
' at zero.  For the above example, you would set the filter index to 5 to
' select "All Files"

    Dim filename As Variant

'   Get the file name
    filename = Application.GetSaveAsFilename _
        (InitialFileName:=defname, _
         FileFilter:=Filt, _
         FilterIndex:=FilterIndex, _
         Title:=Prompt)

'   Check if dialog box canceled
    If filename = False Then
        GetExportFileName = ""
    Else
        GetExportFileName = filename
    End If
   
End Function

Function GetImportFileName(Filt As String, FilterIndex As String, Prompt As String) As String

'Selects a single file for import
    Dim filename As Variant
    
'
'Selects a file for export
'
'Use the following format to set up the filter
'    Filt = "Text Files (*.txt),*.txt," & _
'           "Lotus Files (*.prn),*.prn," & _
'           "Comma Separated Files (*.csv),*.csv," & _
'           "ASCII Files (*.asc),*.asc," & _
'           "XML Files (*.xml),*.xml," & _
'           "All Files (*.*),*.*"
'
' The Filter Index controls what filter is selected by default.  It starts
' at zero.  For the above example, you would set the filter index to 5 to
' select "All Files"


'   Get the file name
    filename = Application.GetOpenFilename _
        (FileFilter:=Filt, _
         FilterIndex:=FilterIndex, _
         Title:=Prompt)

'   Exit if dialog box canceled
    If filename = False Then
        GetImportFileName = ""
    Else
        GetImportFileName = filename
    End If
   
End Function

Sub GetImportFileName2()

'Allows multiple files to be selected

    Dim Filt As String
    Dim FilterIndex As Integer
    Dim filename As Variant
    Dim Title As String
    Dim i As Integer
    Dim Msg As String
    
'   Set up list of file filters
    Filt = "Text Files (*.txt),*.txt," & _
            "Lotus Files (*.prn),*.prn," & _
            "Comma Separated Files (*.csv),*.csv," & _
            "ASCII Files (*.asc),*.asc," & _
            "All Files (*.*),*.*"

'   Display *.* by default
    FilterIndex = 5

'   Set the dialog box caption
    Title = "Select a File to Import"

'   Get the file name
    filename = Application.GetOpenFilename _
        (FileFilter:=Filt, _
         FilterIndex:=FilterIndex, _
         Title:=Title, _
         MultiSelect:=True)

'   Exit if dialog box canceled
    If Not IsArray(filename) Then
        MsgBox "No file was selected."
        Exit Sub
    End If
   
'   Display full path and name of the files
    For i = LBound(filename) To UBound(filename)
        Msg = Msg & filename(i) & vbCrLf
    Next i
    MsgBox "You selected:" & vbCrLf & Msg
End Sub

