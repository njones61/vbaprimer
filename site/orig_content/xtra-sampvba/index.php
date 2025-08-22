<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Excel VBA Primer</title>
<link href="../../nljstyles.css" rel="stylesheet" type="text/css" />
<link href="../primer.css" rel="stylesheet" />
<link href="../../prism/prism.css" rel="stylesheet" />

</head>

<body>
<script src="../../prism/prism.js"></script>

<?php 
require "../header.php";
?>
<h1>Sample VBA Applications</h1>
<p>The following spreadsheets, documents, and VB modules represent examples of 
how VBA can be used in advanced applications within Excel and by other programs 
in addition to Excel.&nbsp; Right click on each link to download.&nbsp; I 
typically discuss these samples at the end of the semester.</p>
<table width="1126" cellspacing="8" >
  <tr>
    <td width="16%" valign="top"><b>Title</b></td>
    <td width="14%" align="center" valign="top"><b>Link</b></td>
    <td width="70%" valign="top"><b>Description</b></td>
  </tr>
  <tr>
    <td width="16%" valign="top">Weekly Appt. Message Generator</td>
    <td width="14%" align="center" valign="top"> <a href="MailAppts.bas">MailAppts.bas</a></td>
    <td width="70%" valign="top">This is a VBA module for 
      Outlook that will search through appointments in the default calendar and 
      find all appointments of a specified type within a given window (one week 
      for example) of the current date.&nbsp; These appointments are then listed 
      in a new html formatted e-mail message.</td>
  </tr>
  <tr>
    <td width="16%" valign="top">MS Word Calendar Generator</td>
    <td width="14%" align="center" valign="top"> <a href="calendar.dot">calendar.dot</a></td>
    <td width="70%" valign="top">This is a macro for MS Word 
      that I found on the internet.&nbsp; It is attached to a document template.&nbsp; 
      Simply open up the template in Word and it should prompt you with a user 
      form.&nbsp; If that doesn't work, go to the Macros command in the Tools menu 
      and manually start the macro.&nbsp; This macro will search through your 
      default calendar in Outlook and generate a beautifully formatted calendar in 
      Word that you can print or save.</td>
  </tr>
  <tr>
    <td width="16%" valign="top">MS Word Remove Line Breaks 
      Macro</td>
    <td width="14%" align="center" valign="top"> <a href="RemoveBreaks.doc">RemoveBreaks.doc</a></td>
    <td width="70%" valign="top">This is a little macro I 
      recorded in Word that is useful for taking text from an e-mail message and 
      removing the extra line breaks.&nbsp; To use it, open up this document and 
      cut and paste the text from your e-mail message to the document.&nbsp; Then 
      run the macro from the Tools menu.&nbsp; The macro searches through the text 
      in the document and gets rid of the extra line breaks in the middle of 
      paragraphs.&nbsp; It assumes that all paragraphs are delineated with double 
      line returns.</td>
  </tr>
  <tr>
    <td width="16%" valign="top">Auto Mail Script</td>
    <td width="14%" align="center" valign="top"> <a href="ReminderEmail.vbs">ReminderEmail.vbs</a></td>
    <td width="70%" valign="top">This is a VB script that can 
      be run simply by clicking on the file or you can launch it on a repeating 
      basis using Windows scheduling.&nbsp; It generates an e-mail message to a 
      list of recipients.&nbsp; To use this, you will need to add it to your 
      Outlook VB code and then modify a few lines of code (which are all clearly 
      marked).</td>
  </tr>
  <tr>
    <td width="16%" valign="top">AutoCAD Polygon Area Plotter</td>
    <td width="14%" align="center" valign="top"> <a href="polyarea.bas">polyarea.bas</a></td>
    <td width="70%" valign="top">This was sent to me by a 
      former student.&nbsp; I have not tried it personally.&nbsp; It generates a 
      text tag on a set of polygons indicating the area of each polygon.</td>
  </tr>
  <tr>
    <td width="16%" valign="top">Get Directory</td>
    <td width="14%" align="center" valign="top"> <a href="get_dir.bas">get_dir.bas</a></td>
    <td width="70%" valign="top">This is a VB module with 
      functions for prompting the user with the standard Windows dialog for 
      selecting a directory.&nbsp; Can be used in any VBA application.</td>
  </tr>
  <tr>
    <td width="16%" valign="top">Get File</td>
    <td width="14%" align="center" valign="top"> <a href="get_file.bas">get_file.bas</a></td>
    <td width="70%" valign="top">This is a VB module with 
      functions for prompting the user with the standard Windows dialog for 
      selecting a file.&nbsp; Both the import and export version of the file 
      selector dialog are supported.&nbsp; Can be used with any VBA application.</td>
  </tr>
  <tr>
    <td width="16%" valign="top">New Letter</td>
    <td width="14%" align="center" valign="top"> <a href="NewLetter.bas">NewLetter.bas</a></td>
    <td width="70%" valign="top">This is an Outlook macro.&nbsp; 
      If you first select a contact and then run this macro, it will generate a 
      formatted business letter in Word to the contact.</td>
  </tr>
</table>
<?php 
require "../footer.php";
?>
</body>
</html>
