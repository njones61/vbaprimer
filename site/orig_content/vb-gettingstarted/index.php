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
<h1>Getting Started - Using the Developer Interface and the VB Editor</h1>
<p>Visual Basic (VB) is a great programming language for beginning programmers. It has a simple structure and it provides a number of safeguards that prevent common programming errors. Another great feature of VB is that it can be used as  a powerful scripting
  language for writing macros and extensions to Microsoft Office applications including <b>Excel</b> and <strong>Access</strong>. It can also be used to write scripts for use in <strong>AutoCAD</strong>. There is a special version of VB used in these applications called <b> Visual Basic for Applications
(VBA)</b>. </p>
<p>Writing VBA code for Excel  is easy and fun!! Once you learn a few
  basics, you will be creating highly professional spreadsheets. VBA allows
  you to design a spreadsheet that will do things that are impossible with the
  basic spreadsheet options. It also allows you to make your spreadsheets
  more user-friendly.</p>
<h2>The Developer Tab</h2>
<p>The first step in adding Visual Basic to your spreadsheet is to turn on the <b> Developer</b> tab. This is not a default part of the ribbon, so you may need to turn it on as follows:</p>
<ol>
  <li>Select the <strong>File|Options</strong> men  command. </li>
  <li>Click on the <strong>Customize Ribbon</strong> button on the left.</li>
  <li>Turn on the <strong>Developer</strong> option shown in the <strong>Main Tabs</strong> section on the right.</li>
</ol>
<blockquote>
  <p><img src="exceloptions.png" width="840" height="685" /></p>
</blockquote>
<ol start="4">
  <li>Click <strong>OK</strong> to exit.</li>
</ol>
<p> You should now see the <strong>Developer</strong> tab. This is where we interact with our VB code.</p>
<p><img src="devtab.png" width="853" height="219" /></p>
<h2>The Code Group</h2>
<p>The <strong>Code</strong> group is used to record <a href="../vb-macros/index.php">macros</a> and to open the VB editor. The <strong>Visual Basic</strong> button opens the Visual Basic Editor window and the other tools are used to record and control macros.</p>
<p><img src="codegroup.png" width="237" height="88" /></p>
<h2>The Visual Basic Editor</h2>
<p>The <b>VB Editor</b> is where you edit the Visual Basic code.&nbsp; It is
  very similar to the regular Visual Basic compiler.&nbsp; The code is shown in a
  set of windows on the right.&nbsp; The <b>Project</b> window on the left lists
  the components of the project.&nbsp; The <b>VBAProject</b> folder
lists each of the sheets in your spreadsheet and the workbook.&nbsp; The <b>Modules</b> folder lists the code associated with <a href="#Using Macros">Macros</a>. The Forms folder lists the custom user forms associated with the project. To edit the code associated with a sheet, module, or user form, you simply double-click on the object in the Project Explorer Window.</p>
<p><img src="vbeditor.png" width="1003" height="897" /></p>
<h2>The Controls Group</h2>
<p>The <strong>Controls</strong> group is used to add <a href="../vb-controls/index.php">controls</a> to a worksheet and to create/edit the VB code associated with the controls. </p>
<p><img src="controlsgroup.png" width="172" height="94" /></p>
<p>The <strong>View Code</strong> button brings up the Visual Basic Editor window shown above.</p>
<h2>Security Settings</h2>
<p>Since VBA is such a flexible and powerful scripting environment, it also happens to be a popular method for writing viruses. For example, it is possible to write scripts that are automatically executed whenever a spreadsheet is opened. The script could theoretically attempt to do some damage to your computer (delete files, etc.) once it executes. To minimize the chance that a malicious script could cause damage, Microsoft turns on some default layers of security over VBA scripts. Before we can start writing VBA code, we need to adjust those settings.</p>
<ol>
  <li>Go to the Developer tab.</li>
  <li>Click on Macro Security</li>
</ol>
<p>You will then be presented with the following options:</p>
<p><img src="security.png" width="672" height="444" alt=""/></p>
<p>Select the settings shown in the figure above and click OK.</p>
<p>You should only need to do this once. These settings are associated with your installation of Excel and will be applied each time you open a spreadsheet from here on out.</p>
<?php 
require "../footer.php";
?>
</body>
</html>
