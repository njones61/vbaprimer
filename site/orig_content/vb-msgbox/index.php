<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Excel VBA Primer</title>
<link href="../../nljstyles.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<!--
.style1 {font-family: "Courier New", Courier, monospace}
-->
</style>
<link href="../primer.css" rel="stylesheet" />
<link href="../../prism/prism.css" rel="stylesheet" />

</head>

<body>
<script src="../../prism/prism.js"></script>

<?php 
require "../header.php";
?>
<h1>MsgBox and InputBox</h1>
<p>There are many occasions when programming where it is convenient to show a prompt to display a message to the user, and in some cases ask a simple question or ask for some simple feedback. In VBA, this can be accomplished with the MsgBox and InputBox tools.</p>
<h2>MsgBox</h2>
<p>MsgBox is used to display a simple message to the user. Here is the syntax:</p>
<blockquote>
  <p><img src="syntax_msgbox.png" width="588" height="17" alt=""/></p>
</blockquote>
<p>The last two arguments (HelpFile, Context) are rarely used. The other arguments are described below.</p>
<h3>Prompt</h3>
<p>Note that the only argument that is required is the Prompt. Here is a simple example. The code:</p>

<pre><code class="language-vb">MsgBox "Hello World"
</code></pre>

<p>Brings up the following window:</p>
<blockquote>
  <p><img src="sample_msgbox_helloworld.png" width="154" height="154" alt=""/></p>
</blockquote>
<h3>Buttons</h3>
<p>The second argument can be used to control what buttons are displayed. For example, the code:</p>

<pre><code class="language-vb">MsgBox "Hello World", vbOKCancel
</code></pre>

<p>Brings up the following window:</p>
<blockquote>
  <p><img src="sample_msgbox_okcancel.png" width="251" height="154" alt=""/></p>
</blockquote>
<p>The buttons argument is what is called an enumerated type, meaning you can only pick from a pre-defined set of options. The options are defined as VB constants, hence the &quot;vb&quot; prefix. Here are the more commonly used options:</p>
<blockquote>
  <p>vbOKOnly<br />
    vbOKCancel<br />
    vbRetryCancel
    <br />
  vbYesNo<br />
  vbYesNoCancel
  </p>
</blockquote>
<p>To determine which button is selected, you need to apply MsgBox as a function rather than as a sub. For example:</p>

<pre><code class="language-vb">Dim mybutton As Variant
mybutton = MsgBox("This will delete your sheet. Continue?", vbOKCancel)
If mybutton = vbOK Then
    'PUT THE CODE HERE TO DELETE THE SHEET
Else
    Exit Sub
End If
</code></pre>

<p>In other words, when used as a function, MsgBox returns a code indicating the button that was selected. Once again, the button codes are defined as a set of VB constants:</p>
<blockquote>
  <p>vbOK<br />
    vbCancel<br />
    vbYes<br />
    vbNo<br />
  vbRetry</p>
</blockquote>
<h3>Style</h3>
<p>You can also choose to add an icon to the MsgBox by adding another constant to the button argument. For example, the code:</p>

<pre><code class="language-vb">MsgBox "This sheet brought to you by Norm Jones", vbInformation + vbOKOnly
</code></pre>


<p>Brings up the following window:</p>
<blockquote>
  <p><img src="sample_msgbox_stylebutton.png" width="336" height="171" alt=""/></p>
</blockquote>
<p>Note that the style constants can be added to the button constants. The style constants are as follows:</p>
<blockquote>
  <table width="200" border="0" cellpadding="4">
    <tr>
      <td align="center"><img src="msgbox_style_information.png" width="32" height="32" alt=""/></td>
      <td align="left">vbInformation</td>
    </tr>
    <tr>
      <td align="center"><img src="msgbox_style_exclarmation.png" width="31" height="28" alt=""/></td>
      <td align="left">vbExclamation</td>
    </tr>
    <tr>
      <td align="center"><img src="msgbox_style_critical.png" width="32" height="32" alt=""/></td>
      <td align="left">vbCritical</td>
    </tr>
  </table>
</blockquote>
<h3>Title</h3>
<p>For each of the examples show above, note that the text shown in the title bar is &quot;Microsoft Excel&quot;. You can change the title by using the third argument as follows:</p>

<pre><code class="language-vb">MsgBox "Hello World", vbOKOnly, "Greetings"
</code></pre>


<p>This code brings up:</p>
<blockquote>
  <p><img src="sample_msgbox_title.png" width="154" height="154" alt=""/></p>
</blockquote>
<h1>InputBox</h1>
<p>InputBox is very similar to MsgBox, but it is used when you need to prompt the user to input some text (or a number or a date) before you execute some code. The syntax is as follows:</p>
<blockquote>
  <p><img src="syntax_inputbox.png" width="442" height="17" alt=""/></p>
</blockquote>
<h3>Prompt</h3>
<p>Once again, the prompt is the message that is displayed and it is the only required argument. For example, the code:</p>

<pre><code class="language-vb">name = inputbox("Please enter your name")
</code></pre>


<p>brings up:</p>
<blockquote>
  <p><img src="sample_inputbox_promptonly.png" width="373" height="158" alt=""/></p>
</blockquote>
<p>Note that InputBox is always used as a function. The value returned by the function is the text string entered by the user. If the user selects the Cancel button, InputBox returns an empty string. Therefore, to determine what button was selected, you simply test the value of the return string as follows:</p>

<pre><code class="language-vb">result = inputbox("Please enter your name")
If result <> "" Then
    'DO SOMETHING WITH NAME HERE
End If
</code></pre>


<h3>Title</h3>
<p>The <strong>Title</strong> argument is used to specify a text string to go in the title bar, similar to MsgBox.</p>
<h3>Default</h3>
<p>The Default argument is used to provide a default text string in the input box when it first comes up. For example,</p>

<pre><code class="language-vb">result = inputbox("Please enter your name", "Greetings", "Joe Blow")
If result <> "" Then
    'DO SOMETHING WITH NAME HERE
End If
</code></pre>


<p>brings up:</p>
<blockquote>
  <p><img src="syntax_inputbox2.png" width="373" height="158" alt=""/></p>
</blockquote>

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
      <td> <strong>MsgBox -</strong> Create a control button that displays a message box with a congratulatory statement.</td>
      <td align="center" valign="top">Easy</td>
      <td align="center" valign="top"><a href="msgbox.xlsm">msgbox.xlsm</a></td>
      <td align="center" valign="top"><a href="msgbox_key.xlsm">msgbox_key.xlsm</a></td>
    </tr>
    <tr>
      <td><strong>Unit Weight - </strong>Create a message box that will only calculate the unit weight if the user clicks on "OK."</td>
      <td align="center" valign="top">Medium</td>
      <td align="center" valign="top"><a href="unit_weight.xlsm">unit_weight.xlsm</a></td>
      <td align="center" valign="top"><a href="unit_weight_key.xlsm">unit_weight_key.xlsm</a></td>
    </tr>
    <tr>
      <td><strong>Scholarship Letter -</strong> Create an input box that asks for the recipient's name and adds it to the scholarship award letter template.</td>
      <td align="center" valign="top">Hard</td>
      <td align="center" valign="top"><a href="scholarship_letter.xlsm">scholarship_letter.xlsm</a></td>
      <td align="center" valign="top"><a href="scholarship_letter_key.xlsm">scholarship_letter_key.xlsm</a></td>
    </tr>
  </tbody>
</table>
<br />

<?php 
require "../footer.php";
?>
</body>
</html>
