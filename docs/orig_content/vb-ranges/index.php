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
<h1>Working with Cells and Ranges</h1>
<p>When writing VB code, you can use variables, for loops, and all other VB
  types and statements.&nbsp; However, most of your code will be dealing with
  values stored in cells and ranges.</p>
<h2>The Range Object</h2>
<p>A <b>range</b> is set of cells.&nbsp;
  It can be one cell or multiple cells.&nbsp; A range is an object within a
  worksheet object.&nbsp; For example, the following statement sets the value of
  
  cell C23 to a formula referencing a set of named cells:</p>

<pre><code class="language-vb">Worksheets("trapchan").Range("C23").Value = "=(u/n)*A*Rh^(2/3)*S^(1/2)"
</code></pre>  
  

<p>If the &quot;trapchan&quot; worksheet is the active sheet, the first part can
  be left off as follows:</p>
  
<pre><code class="language-vb">Range("C23").Value = "=(u/n)*A*Rh^(2/3)*S^(1/2)"
</code></pre>  
  

<p>The &quot;.Value&quot; part is optional.&nbsp; You 
  can also write:</p>

<pre><code class="language-vb">Range("C23") = "=(u/n)*A*Rh^(2/3)*S^(1/2)"
</code></pre>
  

<p>You have to remember to put the double quotes around the cell address.&nbsp;
  If the cell C23 has been named &quot;Q&quot; in the spreadsheet, you can
  reference the range as follows:</p>
  
<pre><code class="language-vb">Range("Q").Value = "=(u/n)*A*Rh^(2/3)*S^(1/2)"
</code></pre>
  

<p>To get something from a cell and put it in a variable, you just do things in 
  reverse:</p>
  
<pre><code class="language-vb">x = Range("B14").Value
</code></pre>


<h2>The Cells Object</h2>
<p>Another way to interact with a cell is to use the <b> Cells(rowindex, columnindex)</b> function.&nbsp; For example:</p>

<pre><code class="language-vb">Cells(2, 5).Value = "=(u/n)*A*Rh^(2/3)*S^(1/2)"
</code></pre>


<p>or</p>

<pre><code class="language-vb">x = val(Cells(2, 14).Value)
</code></pre>


<p>Once again, the value part is optional since it is the default property of 
  both the cell and range object.</p>
<h2>Working with Multiple Cells</h2>
<p>A range can also encompass a set of cells.&nbsp; The following code selects a
  block of cells:</p>

<pre><code class="language-vb">Range("A1:C5").Select
</code></pre>
  
<p>or</p>

<pre><code class="language-vb">Range("A1", "C5").Select
</code></pre>

<p>In some cases, it is useful to reference a range of cells using
  integers representing the row and column of the cell.&nbsp; This can be
  accomplished with the <b>Cells</b> object.&nbsp; For example:</p>
  
<pre><code class="language-vb">Range(Cells(1,1), Cells(3,5)).Select
</code></pre>

<p>The problem with referring to specific cells in your code is that if you 
  change the location of data on your spreadsheet, you need to go through your 
  code and make sure all of the addresses are updated.&nbsp; In some cases it it 
  useful to define the ranges you are dealing with using global constants at the 
  top of your VB code.&nbsp; Then when you reference a range, you can use the 
  constants.&nbsp; If the range ever changes, you only need to update your code in 
  one location.&nbsp; For example:</p>
  
<pre><code class="language-vb">Const TableRange As String = "A4:D50"
Const NameRange As String = "A4:D50"
Const ScoreRange As String = "D4:D50"

.
.
Range(TableRange).ClearContents
.
.
</code></pre>
  

<p>An even better approach is to get into the habit of naming cells and ranges 
  on your spreadsheet.&nbsp; Then your VB code can always refer to ranges by names.&nbsp; 
  Then, if you change the location or domain of a named range, you generally don't 
  need to update your VB code.&nbsp; For example:</p>

<pre><code class="language-vb">Range("TableRange").ClearContents
Range("NameRange").ClearContents
Range("ScoreRange").ClearContents
</code></pre>
  

<h2>Looping Through Cells</h2>
<p>One of the most common things we do with VB code is to traverse or loop 
  through a set of cells.&nbsp; There are several ways this can be accomplished.&nbsp; 
  One way is to use the <b>Cells</b> object.&nbsp; The following code loops 
  through a table of cells located in the range B4:F20:</p>
  
<pre><code class="language-vb">Dim row As Integer
Dim col As Integer

For row = 4 To 20
   For col = 2 To 6
      If Cells(row, col) = "" Then
         Cells(row, col).Interior.Color = vbRed
      End If
   Next col
Next row
</code></pre>
  

<p>In most cases, it doesn't matter what order the cells are traversed, as long 
  as each cell is visited.&nbsp; In such cases the <b>For Each ... Next</b> looping style may be used.&nbsp; The following code does the same thing as the 
  nested loop shown above:</p>
  
<pre><code class="language-vb">Dim cell As Variant

For Each cell In Range("B4:F20")
   If cell = "" Then
      cell.Interior.Color = vbRed
   End If
Next cell
</code></pre>

<p>Another option is to create your own range objects.&nbsp; A
  range object is essentially a range type variable.&nbsp; The following code
  defines three ranges:</p>
  
<pre><code class="language-vb">Dim coordrange As Range
Dim xrange As Range
Dim yrange As Range

Set xrange = Range(Cells(3, 1), Cells(100, 1))
Set yrange = Range(Cells(3, 2), Cells(100, 2))
Set coordrange = Range(Cells(3, 1), Cells(100, 2))
</code></pre>
  
<p>Once a set of range objects has been defined, you can easily manipulate the
  cells in the range object.&nbsp; For example, the following code clears the
  contents of all the cells in the coordrange object:</p>

<pre><code class="language-vb">coordrange.Clear
</code></pre>
  

<p>Once again, these ranges can be traversed using the <b>For Each ... Next</b> syntax.</p>

<pre><code class="language-vb">Dim cell As Variant

For Each cell In xrange
    cell.Value = "0.0"
Next cell
</code></pre>


<p>One of the most useful things you can do with VBA in Excel is to allow the
  user to enter a list of numbers where the size of the list can vary.&nbsp; The
  following code searches through a list and copies the numbers in the list into
  an array.&nbsp; It stops copying the numbers when it reaches a blank cell.</p>
  
<pre><code class="language-vb">'Get the x values
i = 0
For Each cell In Range("B4:B23")
    If cell.Value = "" Then
        numpts = i
        Exit For
    Else
        i = i + 1
        x(i) = cell.Value
    End If
Next cell
</code></pre>
  

<h2>R1C1 Notation</h2>
<p>Another case where we deal with ranges in VB code is when we record macros and then modify the macros to make custom subs. If we create a formula while recording the macro, the formula will often be recorded using what is called &quot;R1C1 notation&quot;. For example, here is a snippet of code from a recorded macro:</p>

<pre><code class="language-vb">Range("F11").Select
ActiveCell.FormulaR1C1 = "=SUM(R[-4]C:R[-2]C)"
</code></pre>

<p>This is the same as:</p>

<pre><code class="language-vb">Range("F11").Select
ActiveCell.Formula = "=SUM(F7:F9)"
</code></pre>


<p>So what is the advantage of the R1C1 notation? If you need to modify the formula using variables or some other run-time conditions, the R1C1 notation can be easier to manipulate. The R1C1 notation makes it possible to refer to both rows and columns using integers rather than letters and it makes it very easy to indicate relative vs. absolute references. The following table illustrates how the R1C1 notation works:</p>
<table width="358" border="1">
  <tr>
    <td width="164" align="center"><strong>R1C1 Expression</strong></td>
    <td width="184" align="center"><strong>Equivalent Expression*</strong></td>
  </tr>
  <tr>
    <td align="center"><span class="code">R[-4]C:R[-2]C</span></td>
    <td align="center">F7:F9</td>
  </tr>
  <tr>
    <td align="center"><span class="code">R[+2]C:R[+4]C</span></td>
    <td width="1" height="1" align="center">F13:F15</td>
  </tr>
  <tr>
    <td align="center"><span class="code">R4C2:RC[-1]</span></td>
    <td align="center">$B$4:E11</td>
  </tr>
  <tr>
    <td align="center"><span class="code">RC:R15C10</span></td>
    <td align="center">F11:$J$15</td>
  </tr>
</table>
*Assumes reference cell = F11
<p>In other words, R10 means &quot;an absolute reference to row 10&quot; whereas R[-2] means &quot;a relative reference to two rows above the cell in question&quot; and R (without a number) means &quot;on the same row as the cell in question&quot;.</p>

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
      <td> <strong>Range Basics -</strong> Run through 5 quick exercises that use ranges.</td>
      <td align="center" valign="top">Easy</td>
      <td align="center" valign="top"><a href="range_basics.xlsm">range_basics.xlsm</a></td>
      <td align="center" valign="top"><a href="range_basics_key.xlsm">range_basics_key.xlsm</a></td>
    </tr>
    <tr>
      <td><strong>Calculator - </strong>Calculate the given equations from three variables.</td>
      <td align="center" valign="top">Medium</td>
      <td align="center" valign="top"><a href="calculator.xlsm">calculator.xlsm</a></td>
      <td align="center" valign="top"><a href="calculator_key.xlsm">calculator_key.xlsm</a></td>
    </tr>
    <tr>
      <td><strong>Sum - </strong>Use ranges to calculate a sum and an adjusted sum of the tabulated data.</td>
      <td align="center" valign="top">Medium</td>
      <td align="center" valign="top"><a href="sum.xlsm">sum.xlsm</a></td>
      <td align="center" valign="top"><a href="sum_key.xlsm">sum_key.xlsm</a></td>
    </tr>
    <tr>
      <td><strong>Empty Cells -</strong> Calculate the number of empty cells from the range of data. Also change the range object properties to display a red cell for an empty cell.</td>
      <td align="center" valign="top">Hard</td>
      <td align="center" valign="top"><a href="empty_cells.xlsm">empty_cells.xlsm</a></td>
      <td align="center" valign="top"><a href="empty_cells_key.xlsm">empty_cells_key.xlsm</a></td>
    </tr>
  </tbody>
</table>
<br />

<?php 
require "../footer.php";
?>
</body>
</html>
