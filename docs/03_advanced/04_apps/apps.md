# Sample VBA Applications

The following spreadsheets, documents, and VB modules represent examples of how VBA can be used in advanced applications within Excel and by other programs in addition to Excel. Right click on each link to download. I typically discuss these samples at the end of the semester.

| Title | Link | Description |
|-------|------|-------------|
| **Weekly Appt. Message Generator** | [MailAppts.bas](files/MailAppts.bas) | This is a VBA module for Outlook that will search through appointments in the default calendar and find all appointments of a specified type within a given window (one week for example) of the current date. These appointments are then listed in a new html formatted e-mail message. |
| **MS Word Calendar Generator** | [calendar.dot](files/calendar.dot) | This is a macro for MS Word that I found on the internet. It is attached to a document template. Simply open up the template in Word and it should prompt you with a user form. If that doesn't work, go to the Macros command in the Tools menu and manually start the macro. This macro will search through your default calendar in Outlook and generate a beautifully formatted calendar in Word that you can print or save. |
| **MS Word Remove Line Breaks Macro** | [RemoveBreaks.doc](files/RemoveBreaks.doc) | This is a little macro I recorded in Word that is useful for taking text from an e-mail message and removing the extra line breaks. To use it, open up this document and cut and paste the text from your e-mail message to the document. Then run the macro from the Tools menu. The macro searches through the text in the document and gets rid of the extra line breaks in the middle of paragraphs. It assumes that all paragraphs are delineated with double line returns. |
| **Auto Mail Script** | [ReminderEmail.vbs](files/ReminderEmail.vbs) | This is a VB script that can be run simply by clicking on the file or you can launch it on a repeating basis using Windows scheduling. It generates an e-mail message to a list of recipients. To use this, you will need to add it to your Outlook VB code and then modify a few lines of code (which are all clearly marked). |
| **AutoCAD Polygon Area Plotter** | [polyarea.bas](files/polyarea.bas) | This was sent to me by a former student. I have not tried it personally. It generates a text tag on a set of polygons indicating the area of each polygon. |
| **Get Directory** | [get_dir.bas](files/get_dir.bas) | This is a VB module with functions for prompting the user with the standard Windows dialog for selecting a directory. Can be used in any VBA application. |
| **Get File** | [get_file.bas](files/get_file.bas) | This is a VB module with functions for prompting the user with the standard Windows dialog for selecting a file. Both the import and export version of the file selector dialog are supported. Can be used with any VBA application. |
| **New Letter** | [NewLetter.bas](files/NewLetter.bas) | This is an Outlook macro. If you first select a contact and then run this macro, it will generate a formatted business letter in Word to the contact. |

## Application Categories

These sample applications demonstrate several key areas where VBA can be applied:

### Office Integration
- **MailAppts.bas** - Outlook calendar integration
- **NewLetter.bas** - Outlook and Word integration
- **calendar.dot** - Word template with VBA macros

### File System Operations
- **get_dir.bas** - Directory selection dialogs
- **get_file.bas** - File selection dialogs

### Text Processing
- **RemoveBreaks.doc** - Text formatting automation

### Automation Scripts
- **ReminderEmail.vbs** - Automated email generation
- **polyarea.bas** - AutoCAD automation

## How to Use These Samples

### For Excel VBA Developers

1. **Download the .bas files** - These contain VBA code that can be imported into Excel
2. **Import into your project** - Use the Import File command in the VBE
3. **Modify for your needs** - Adapt the code to your specific requirements
4. **Test thoroughly** - Make sure the code works in your environment

### For Word VBA Developers

1. **Download the .dot and .doc files** - These contain Word macros
2. **Open in Word** - The macros should be available automatically
3. **Run from Tools menu** - Use the Macros command to execute

### For Outlook VBA Developers

1. **Download the .bas files** - These contain Outlook VBA code
2. **Import into Outlook VBE** - Use the Visual Basic Editor in Outlook
3. **Modify as needed** - Customize for your specific use case

## Code Examples

Here are some key code snippets from these applications:

### Directory Selection (from get_dir.bas)

```vb
Function GetDirectory(Optional Title As String = "Select Directory") As String
    Dim FSO As Object
    Dim Folder As Object
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set Folder = Application.FileDialog(msoFileDialogFolderPicker)
    
    With Folder
        .Title = Title
        .ButtonName = "Select"
        If .Show = -1 Then
            GetDirectory = .SelectedItems(1)
        Else
            GetDirectory = ""
        End If
    End With
    
    Set FSO = Nothing
    Set Folder = Nothing
End Function
```

### File Selection (from get_file.bas)

```vb
Function GetFile(Optional Title As String = "Select File", _
                Optional Filter As String = "All Files (*.*)|*.*") As String
    Dim FileDialog As Object
    
    Set FileDialog = Application.FileDialog(msoFileDialogFilePicker)
    
    With FileDialog
        .Title = Title
        .Filters.Clear
        .Filters.Add "All Files", "*.*"
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            GetFile = .SelectedItems(1)
        Else
            GetFile = ""
        End If
    End With
    
    Set FileDialog = Nothing
End Function
```

### Email Generation (from MailAppts.bas)

```vb
Sub GenerateAppointmentEmail()
    Dim OutlookApp As Object
    Dim MailItem As Object
    Dim Calendar As Object
    Dim Appointment As Object
    Dim EmailBody As String
    
    ' Create Outlook application object
    Set OutlookApp = CreateObject("Outlook.Application")
    
    ' Create new mail item
    Set MailItem = OutlookApp.CreateItem(0)
    
    ' Set email properties
    With MailItem
        .Subject = "Weekly Appointments"
        .To = "recipient@example.com"
        .HTMLBody = GenerateAppointmentHTML()
        .Display
    End With
    
    Set MailItem = Nothing
    Set OutlookApp = Nothing
End Sub
```

## Best Practices Demonstrated

These sample applications showcase several VBA best practices:

1. **Error Handling** - Many include proper error handling
2. **Object Cleanup** - Proper disposal of COM objects
3. **User Interface** - Standard Windows dialogs for file/directory selection
4. **Modularity** - Functions that can be reused across applications
5. **Documentation** - Clear comments explaining functionality

## Extending These Applications

You can extend these applications by:

- **Adding more features** - Enhance functionality based on your needs
- **Improving error handling** - Add more robust error checking
- **Creating user forms** - Add custom interfaces for better user experience
- **Integrating with databases** - Connect to external data sources
- **Adding logging** - Track operations for debugging and auditing

## Troubleshooting

Common issues when using these samples:

1. **Security settings** - Ensure macros are enabled
2. **References** - Check that required libraries are referenced
3. **Permissions** - Some operations require elevated permissions
4. **Version compatibility** - Test with your specific Office version

These sample applications provide a solid foundation for understanding how VBA can be used across different Office applications and can serve as starting points for your own automation projects.
