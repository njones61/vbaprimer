Attribute VB_Name = "get_directory"
'Option Explicit

'32-bit API declarations
Declare Function SHGetPathFromIDList Lib "shell32.dll" _
  Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long

Declare Function SHBrowseForFolder Lib "shell32.dll" _
Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long

Public Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type


Function GetDirectory1(Msg As String) As String
'   For Excel 97 or later
    Dim bInfo As BROWSEINFO
    Dim path As String
    Dim r As Long, x As Long, pos As Integer
 
'   Root folder = Desktop
    bInfo.pidlRoot = 0&

'   Title in the dialog
    bInfo.lpszTitle = Msg
    
'   Type of directory to return
    bInfo.ulFlags = &H1

'   Display the dialog
    x = SHBrowseForFolder(bInfo)
    
'   Parse the result
    path = Space$(512)
    r = SHGetPathFromIDList(ByVal x, ByVal path)
    If r Then
        pos = InStr(path, Chr$(0))
        GetDirectory1 = Left(path, pos - 1) & "\"
    Else
        GetDirectory1 = ""
    End If
End Function


Function GetDirectory2(Msg As String, defpath As String) As String
'   For Excel 2002
    With Application.FileDialog(msoFileDialogFolderPicker)
'        .InitialFileName = Application.DefaultFilePath & "\"
        .InitialFileName = defpath
        .Title = Msg
        .Show
        If .SelectedItems.Count = 0 Then
            GetDirectory2 = ""
        Else
            GetDirectory2 = .SelectedItems(1) & "\"
        End If
    End With
End Function


Function GetDirectory(Msg As String, defpath As String) As String
'General function that uses appropriate version of the dialog
'Returns empty string if the user selects cancel
    
If Val(Application.Version) < 10 Then
    GetDirectory = GetDirectory1(Msg)
Else
    GetDirectory = GetDirectory2(Msg, defpath)
End If
    
End Function


