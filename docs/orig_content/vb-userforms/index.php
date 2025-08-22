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
<h1>Creating Custom User Forms</h1>
<p>Occasionally it is usefull to prompt the user with a user form and wait for the user to enter some information to the user form before proceeding to execute some code. In many cases, this can be accomplished with <a href="../vb-msgbox/index.php">MsgBox/InputBox</a>, but there are other cases where we need more detailed information and it is not convenient to pull that information from the cells on a sheet. In such cases, it is possible to create a custom user form containing a set of controls. </p>
<p>To illustrate the process of using a custom user form, assume that we have a workbook where we need to routinely add new sheets to the workbook, but rather than adding blank sheets, we want the sheets to be copied from one of three pre-defined templates (stored as hidden sheets). You can download a copy of the spreadsheet used in this example by clicking <a href="sample code.xlsm">here</a>.</p>
<h2>Creating the Form</h2>
<p> The first step is to create the user form. To do this, go the VB Editor and right-click in the empty space in the project explorer window and select Insert|UserForm.</p>
<blockquote>
  <p><img src="add_user_form.png" width="397" height="417" alt=""/></p>
</blockquote>
<p>This will create a new empty user form and launch the control toolbox:</p>
<blockquote>
  <p><img src="new_user_form.png" width="615" height="290" alt=""/></p>
</blockquote>
<p>Next, we will change the name of the user form and change the user form title (default = &quot;UserForm1&quot;). We do this by clicking on the Properties button <img src="propertiesbutton.png" width="17" height="14" alt=""/> in the VB Editor menu. This will dock the Properties window just below the project explorer window. Change the name to &quot;frmAddSheet&quot;. The prefix &quot;frm&quot; identifies the object in code as a user form and the &quot;AddSheet&quot; part makes the name descriptive of the user form's objective. Next, change the Caption property to &quot;New Sheet Options&quot; (or something similarly descriptive). Note that the caption appears in the form title bar.</p>
<p>Now we are ready to add controls to the sheet. Fortunately, this process is identical to the manner in which <a href="../vb-controls/index.php">controls are added to sheets</a>. Add a set of controls to create a layout as follows. Note that you can use the handles on the sides and corners of the user form to resize the form.</p>
<blockquote>
  <p><img src="user_form.png" width="298" height="291" alt=""/></p>
</blockquote>
<p>The text strings (&quot;Sheet title&quot;, &quot;Template:&quot;) are created with label-type controls. Note that the value property of the checkbox has been set to true because we want this option to be on by default.</p>
<p>At any time while editing the sheet, you can test to see how the sheet behaves by pressing the <strong>Run</strong> button <img src="../vb-debugging/command_run.png" width="9" height="15" alt=""/> in the VB Editor menu. To close the form, click on the red X in the upper right corner.</p>
<h2>Launching the Form</h2>
<p>The next thing we need to do is create some way to launch our form. In this case, we will add a button somewhere on the main sheet of our workbook as follows:</p>
<blockquote>
  <p><img src="add_sheet_button.png" width="125" height="44" alt=""/></p>
</blockquote>
<p>Note the use of the ellipsis on the button title (...). This should be added to a button caption whenever the button is used to bring up user form. This is standard user interface protocol. To make the button launch the form, double-click on the button in design mode and type the code for the click event as follows:</p>

<pre><code class="language-vb">Private Sub cmdAddSheet_Click()
frmAddSheet.Show
End Sub
</code></pre>

<p>We are calling the <strong>.Show</strong> method which launches the sheet. Exit design mode and test your button.</p>
<h2>Writing the User Form Code</h2>
<p>Finally, we will write the code associated with the user form. While it is possible to write code associated with the click event for each control, we don't need to do anything while the user is interacting with the controls on the main part of the form. The only controls that need code are the OK and Cancel buttons. For the Cancel button, we will simply write code to make the form go away. For the OK button, we will write code that makes the form go away and then it will use the selections in the form to guide the creation of a new sheet from a template.</p>
<p>For the Cancel button, go to the VB Editor and double-click on the Cancel button. This will bring up the code window for the form. Type the following:</p>

<pre><code class="language-vb">Private Sub cmdCancel_Click()
frmAddSheet.Hide
End Sub
</code></pre>

<p>This executes the <strong>.Hide</strong> method to make the form go away.</p>
<p>Now we need to return to the form editor window. To make it appear again, click on the second icon  <img src="controlwindowbutton.png" width="15" height="13" alt=""/> at the top of the Project Explorer. You can switch back to the code window by clicking on the first icon <img src="codewindowbutton.png" width="14" height="12" alt=""/>. You can also make the control view appear by double-clicking on the form in the Project Explore.</p>
<p>To write the code associated with the OK button, double-click on the OK button and write the following. Note how the control names are used when the sheet is copied. (Note: this assumes that you have three extra sheets named &quot;Template-CI&quot;, &quot;Template-CO&quot;, and &quot;Template-WL&quot; and that the &quot;Add Sheet...&quot; button is on a sheet named &quot;Main&quot;).</p>

<pre><code class="language-vb">Private Sub cmdOK_Click()

Dim template As String

'Hide the sheet
frmAddSheet.Hide

'Determine which template was selected.
If optCustomerInvoice Then
    template = "Template-CI"
ElseIf optChangeOrder Then
    template = "Template-CO"
Else 'Work log
    template = "Template-WL"
End If

'Copy the template to create a new sheet.
Sheets(template).Select
Sheets(template).Copy After:=Sheets(Sheets.Count)

'Make the sheet visible in case the template is hidden
ActiveSheet.Visible = xlSheetVisible

'Rename the sheet
ActiveSheet.Name = txtSheetTitle

'Bring main sheet back to front if necessary
If chkBringToFront = False Then
    Sheets("Main").Select
End If

End Sub
</code></pre>


<p>That's it! Creating custom user forms is easy and fun. </p>

<h2>Exercises</h2>
<p>You may wish to complete following exercises to  gain practice with and reinforce  the topics covered in this chapter:</p>
<table width="777" border="1">
  <tbody>
    <tr>
      <td width="312"><strong>Description</strong></td>
      <td width="84" align="center"><strong>Difficulty</strong></td>
      <td width="161" align="center"><strong>Start</strong></td>
      <td width="192" align="center"><strong>Solution</strong></td>
    </tr>
    <tr>
      <td> <strong>Battleship -</strong> By creating a user form, the user will be able to begin playing Battleship vs the computer.</td>
      <td align="center" valign="top">Easy</td>
      <td align="center" valign="top"><a href="battleship.xlsm">battleship.xlsm</a></td>
      <td align="center" valign="top"><a href="battleship_key.xlsm">battleship_key.xlsm</a></td>
    </tr>
    <tr>
      <td><strong>Rainfall Log - </strong>Create a user form to log recent rainfall and compute the total rainfall to date.</td>
      <td align="center" valign="top">Medium</td>
      <td align="center" valign="top"><a href="rainfall_log.xlsm">rainfall_log.xlsm</a></td>
      <td align="center" valign="top"><a href="rainfall_log_key.xlsm">rainfall_log_key.xlsm</a></td>
    </tr>
    <tr>
      <td><strong>Recruiting Spreadsheet - </strong>Create a user form that allows the user to insert prospective employees into a recruiting table.</td>
      <td align="center" valign="top">Medium</td>
      <td align="center" valign="top"><a href="recruiting_spreadsheet.xlsm">recruiting_spreadsheet.xlsm</a></td>
      <td align="center" valign="top"><a href="recruiting_spreadsheet_key.xlsm">recruiting_spreadsheet_key.xlsm</a></td>
    </tr>
    <tr>
      <td><strong>Advanced Recruiting -</strong> Create a user form that inserts prospective employees into the recruiting table and highlights those who meet criteria chosen from an additional user from.</td>
      <td align="center" valign="top">Hard</td>
      <td align="center" valign="top"><a href="advanced_recruiting.xlsm">advanced_recruiting.xlsm</a></td>
      <td align="center" valign="top"><a href="advanced_recruiting_key.xlsm">advanced_recruiting_key.xlsm</a></td>
    </tr>
  </tbody>
</table>
<br />

<?php 
require "../footer.php";
?>
</body>
</html>
