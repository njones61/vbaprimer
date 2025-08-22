' ***************************************************
'     Script To Send Outlook Mail Item w/Attachment   *
'                                                     *
'     Danny J. Lesandrini     August 28, 2000         *
'     dan@dea.com  or datafast@home.com               *
'     Dean Evans & Associates, Inc                    *
'     www.dea.com/datafast                            *
'                                                     *
'                                                     *
'   This script will uses COM Automation to access    *
'  the Outlook Object Model in order to create and    *
'  send an Outlook Mail Message.                      *
'                                                     *
'   The user is presented with a series of four       *
'  Input Boxes to collect the information necessary   *
'  to create and send the mail item.                  *
'                                                     *
'   1) Receipiant Address                             *
'          Example:  dan@dea.com                      *
'   2) Message Subject                                *
'          Example:  So, how have you been?           *
'   3) Message Body Text                              *
'          Example:  Blah, blah, blah                 *
'   3) Attachment                                     *
'          Example:  "C:\Autoexec.bat"                *
'                                                     *
'                                                     *
'                                                     *
'   Note that the CreateObject call uses the generic  *
'  reference, "Outlook.Application" with no version   *
'  number.  This assures that the script won't break  *
'  just because the client machine doesn't contain    *
'  the correct version of MS Outlook.  The script     *
'  will break, however, if no version of Outlook is   *
'  found on the client machine                        *
'                                                     *
' ***************************************************

Dim objOutlook
Dim objNameSpace

Dim mItem

Const olMailItem = 0

    Set objOutlook = CreateObject("Outlook.application")
    Set objNameSpace = objOutlook.GetNamespace("MAPI")
    Set mItem = objOutlook.CreateItem(olMailItem)

    mItem.To = "bob@aol.com;fred@aol.com;mary@aol.com"
    mItem.Subject = "Friday check-in reminder"
    mItem.Body = "Don't forget to check in this week.  If we don't write something on your sheet, you haven't checked in.  Thanks."

    mItem.Save
    mItem.Send
  

' ****  Clean up
'
    Set mItem = Nothing
    Set objNameSpace = Nothing
    set objOutlook = Nothing
