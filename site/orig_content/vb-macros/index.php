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
<h1>Recording Macros</h1>
<p>A <b>VB Macro</b> is essentially a VB subroutine that executes a series of VB
  statements. To generate a macro, you must first turn on the <a href="../vb-gettingstarted/index.php#developertab">Developer tab</a>. Macros are then created and managed using the tools in the Code Group on the left side of the tab:</p>
<p><img src="../vb-gettingstarted/controlsgroup.png" width="174" height="87" /></p>
<p>You can automate a series of steps in Excel by recording a macro. As you record a macro, Excel converts each of your steps into lines of VB code that are part of a script that can be replayed at a latter point in time to automatically replicate the steps recorded into the script.</p>
<p>The options are as follows:</p>
<table width="100%" border="0" cellpadding="4" cellspacing="4">
  <tr>
    <td align="center"><img src="recordmacro.png" width="93" height="19" /></td>
    <td width="87%">Click on this button to start recording your macro. You will then be prompted for the name of the macro. Once you start recording the macro, the button will change to Stop Recording.</td>
  </tr>
  <tr>
    <td align="center"><img src="stoprecording.png" width="104" height="16" /></td>
    <td>Click here when you are finished with the steps you wish to include in the macro.</td>
  </tr>
  <tr>
    <td align="center"><img src="userelbutton.png" width="147" height="23" /></td>
    <td>This is a toggle button. When it is turned on, the steps in the macro are recorded in a relative fashion. I.e., the cells affected by the macro will be based on the position of the cell(s) selected when the macro is executed. This can be used to make a macro that can be applied to any section of the spreadsheet. When using this option, be sure to select an appropriate part of the sheet prior to recording the macro.</td>
  </tr>
  <tr>
    <td align="center"><img src="security.png" width="100" height="22" /></td>
    <td width="87%">This option is used to establish the security settings for
      the VB code. VB macros can be used to write computer viruses.&nbsp; The
      security settings are used to minimize danger from such viruses.</td>
  </tr>
  <tr>
    <td align="center"><img src="macrosbutton.png" width="43" height="47" /></td>
    <td>This button brings up a window listing all of the macros associated with a project. You can select a macro and click on the <strong>Run</strong> button to execute the macro. You can also delete macros.</td>
  </tr>
  <tr>
    <td align="center"><img src="vbebutton.png" width="38" height="64" /></td>
    <td width="87%">This tool displays the <a href="../vb-gettingstarted/index.php#vbeditor">Visual Basic
      Editor</a>.&nbsp; This is where you write the Visual Basic code. It also allows you to look at the code associated with your macros. The macros are stored in the <strong>Modules</strong> section of the Project Explorer.</td>
  </tr>
</table>
<p>Macros are extremely useful when you are first learning how to write VBA code
  in Excel. If you want to do something in code such as change the
  background color of a cell, but you don't know to do it, simply run a macro,
  change the color manually, and then look at the macro. You can learn how
to do just about anything simply by running macros.</p>
<h2>Recording a Macro</h2>
<p>To illustrate the macro recording process, open up a blank workbook, click on the Developer tab, and click on the Record Macro button. You will then be prompted for the name of the macro. Enter &quot;my_macro&quot;. Note that you should not use spaces in your macro names.</p>
<p><img src="recordmacrodialog.png" width="358" height="294" alt=""/></p>
<p>Select any cell other than cell B3. Then, do the following:</p>
<ol>
  <li>Select cell <strong>B3</strong>.</li>
  <li>Enter a value of <strong>23.4</strong> in cell <strong>B3</strong>.</li>
  <li>Enter a vlaue of <strong>873.2</strong> in cell <strong>B4</strong>.</li>
  <li>Enter a formula in cell <strong>B5</strong> to compute the sum of the previous two values (&quot;=Sum(B3:B4)&quot;).</li>
  <li>Click on the <strong>Home</strong> tab and select the range <strong>B3:B4</strong> and click on the center align button in the <strong>Alignment</strong> section.</li>
  <li>Use to borders tool in the <strong>Font</strong> section to apply a solid border to the cells.</li>
</ol>
<p>At this point, you workbook should look like this:</p>
<p><img src="macroresults.png" width="260" height="145" alt=""/></p>
<p>Go back to the <strong>Developer</strong> tab and click on the <strong>Stop Recording</strong> button. This completes the recording of the macro.</p>
<h2>Saving a Macro-Enabled Workbook</h2>
<p>Before looking at the source code recorded by our macro, we need to save the changes that we have made into the workbook thus far. Click on the File|Save As... command and pick a location to save your workbook. Note that the filename and filter will look something like this by default:</p>
<p><img src="file-filters-default.png" width="612" height="142" alt=""/></p>
<p>Note that the default extension is &quot;*.xlsx&quot;. If you click the Save button, you will get the following error message:</p>
<p><img src="file-warning.png" width="634" height="179" alt=""/></p>
<p>If we save the file in this format, our macro code will be lost and it will not function the next time we open it. To save a workbook containing VB code, you must change the extension as follows:</p>
<p><img src="file-filters-macro.png" width="609" height="124" alt=""/></p>
<h2>Running a Macro</h2>
<p>Next we will test the macro that we just recorded. Before running the macro, we need to delete our changes to the sheet, including the formatting. The easiest way to do this is to select the entire column <strong>B</strong> and then select the <strong>Delete Sheet Columns</strong> command in the <strong>Cells</strong> section on the right side of the ribbon in the <strong>Home</strong> tab. After the deletion, there should be nothing on the sheet. Select any cell on the sheet. </p>
<p>To run the macro, go to the <strong>Developer</strong> tab and click on the <strong>Macros</strong> button. This will bring up a list of the macros (we only have one at this point). Sure our macro is highlighted and then click on the <strong>Run</strong> button. This should execute our macro and reproduce all of the edits and changes we made to the worksheet when we were recording the macro.</p>
<h2>Viewing the Macro Code</h2>
<p>Next, we will examine the code associated with our macro. When you record a macro in Excel, everything you do is recorded as a set of Visual Basic code. To view the code, click on <strong>Visual Basic</strong> button in the <strong>Developer</strong> tab. This will launch the Visual Basic Editor window. On the left side of the window in the VBA Project Explorer you will see a folder called &quot;Modules&quot;. Click on the plus sign (&quot;+&quot;) to the left of this folder to expand its contents and you will an item called <strong>Module1</strong>. When a macro is recorded, the code associated with the macro is always inserted into a module. If you already have one or more modules, a new module will be created.</p>
<p><img src="module.png" width="235" height="267" alt=""/></p>
<p>If you double click on the item labeled <strong>Module1</strong>, a window will open and the code associated with the macro  you just recorded will appear. The code should look something like this:</p>
<pre><code class="language-vb">Sub my_macro()
'
' my_macro Macro
'

'
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "23.4"
    Range("B4").Select
    ActiveCell.FormulaR1C1 = "873.2"
    Range("B5").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-2]C:R[-1]C)"
    Range("B3:B5").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub
</code></pre>
<p>The first few lines of code are associated with entering the two values and the formula. The next section centers the selection and the last six blocks of code apply the border style to the selection. You can modify the code here if you wish and re-run the macro.</p>
<h2>Applications and Use Cases</h2>
<p>Macros have a number of useful applications. For example, you can record a macro associated with some set of steps that you find yourself doing frequently. The macro then automates those steps. You can also use a macro to figure out how to write VB code to perform some action. For example, suppose I want to embed an execution of the <a href="../excel-goalseek/">Goal Seek</a> feature as part of my code, but I am not sure how to call Goal Seek from VB. I can simply record a macro that involves running the Goal Seek tool and then examine the code.</p>
<p>In may applications, the VB code associated with our macro may not be as useful as we like because it only applies to a very specific case at a very specific location on our workbook. In these cases, we can often modify the macro code to make it more general purpose. For example, we can generalize the code and then put it in a loop to solve some kind of problem in an iterative fashion. In this sense, macros can be a very powerful way to quickly generate code.</p>

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
      <td> <strong>Parobolic Zero -</strong> Record a macro performing a goal seek to find the zero(s) of a parabolic equation.</td>
      <td align="center" valign="top">Easy</td>
      <td align="center" valign="top"><a href="parabolic_zero.xlsm">parabolic_zero.xlsm</a></td>
      <td align="center" valign="top"><a href="parabolic_zero_key.xlsm">parabolic_zero_key.xlsm</a></td>
    </tr>
    <tr>
      <td><strong>Table Format - </strong>Record a relative macro that formats the appropriate range of cells found within any given table of the same size.</td>
      <td align="center" valign="top">Medium</td>
      <td align="center" valign="top"><a href="table_format.xlsm">table_format.xlsm</a></td>
      <td align="center" valign="top"><a href="table_format_key.xlsm">table_format_key.xlsm</a></td>
    </tr>
    <tr>
      <td><strong>Microstrain -</strong> Record a macro that conducts an analysis on microstrain data that quickly identifies the maximum values of the selected column.</td>
      <td align="center" valign="top">Hard</td>
      <td align="center" valign="top"><a href="microstrain.xlsm">microstrain.xlsm</a></td>
      <td align="center" valign="top"><a href="microstrain_key.xlsm">microstrain_key.xlsm</a></td>
    </tr>
  </tbody>
</table>
<br />

<?php 
require "../footer.php";
?>
</body>
</html>
