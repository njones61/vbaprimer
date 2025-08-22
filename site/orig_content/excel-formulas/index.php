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

<h1> Cells and Formulas</h1>

<p>The most basic feature of Excel is the ability to enter data and then write formulas based on the data. As the data are edited, the formulas are automatically updated. In this chapter we review some of the procedure for entering and using formulas. </p>
<h2>Cell Addresses</h2>
<p>An Excel workbook contains a collection of sheets. Each sheet contains a collection of cells organized into rows and columns. The rows are indexed with numbers (1,2,3...) and the columns are indexed with letters (A,B,C...). Each cell can be uniquely identified by a cells address defined by the column-row combination.  </p>
<blockquote>
  <table width="267" border="0">
    <tr>
      <td width="54">A5</td>
      <td width="203">&lt;- Row 5, Column 1</td>
    </tr>
    <tr>
      <td>D3</td>
      <td>&lt;- Row 3, Column 4</td>
    </tr>
    <tr>
      <td>AJ234</td>
      <td>&lt;- Row 234, Column 36</td>
    </tr>
  </table>
</blockquote>
<p>Note that after column Z (26), the column numbers are indexed as AA,AB,AC... To reference a group of cells with single address we combine the upper left corner of the region with the lower right corner separated by a semicolon. For example, to reference the following range:</p>
<p><img src="range.png" width="491" height="267" alt=""/></p>
<p>we would use the address <strong>B3:F10</strong>.</p>
<h2>Cell Inputs</h2>
<p>There are four primary types of information that can be entered in cells:</p>
<ol>
  <li>Text (&quot;Hello world&quot;, etc.)</li>
  <li>Numbers (4, 2382.23, 1e-14, etc.)</li>
  <li>Dates (Jun-5, 2014, 12/29/2015, etc.)</li>
  <li>Formulas (&quot;=A4+C5&quot;, &quot;=Sum(D4:D14)&quot;, etc.)</li>
</ol>
<p>For the first three types (text, numbers, dates), Excel determines the type of data based on the content as you enter it, and formats it appropriately. You can also customize the formatting if you wish. Entering a formula is described in the next section.</p>
<p>Sometimes it is useful to enter a sequence of data in a cell. Excel provides a simple trick for doing this. For example, suppose you want to create a list of numbers 1, 2, 3, ... to fill in a column in a table. Rather than typing the entire list, you can enter the first three numbers and then select the three numbers. Once you do so, a green rectangle will appear at the lower right corner of the selection as follows:</p>
<p><img src="autofill-1.png" width="202" height="219" alt=""/></p>
<p>You can then click on the rectangle handle and drag it all the way down to the bottom of the list, or you can simply double-click on the handle. In either case, the list will automatically be extended as follows:</p>
<p><img src="autofill-2.png" width="220" height="380" alt=""/></p>
<p>This process works for other types of data also. For example, you can enter &quot;Mon&quot;, &quot;Tues&quot;, &quot;Wed&quot; or &quot;Jan&quot;, &quot;Feb&quot;, &quot;Mar&quot; and when you extend the list, the sequences will be automatically extended.</p>
<h2>Entering a Formula</h2>
<p>An Excel formula typically references cells in your worksheet and performs some type of calculation. To enter a formula, you start by typing an equal sign (&quot;=&quot;) and then you reference cells using their addresses (&quot;A4&quot;, &quot;C27&quot;, &quot;D4:E15&quot;, etc.). The values of the cells references are then used in the formula and a value is returned and displayed in the cell. As you change the values of the input cells, all of the dependent formulas are automatically updated. Formulas can reference other cells that contain formulas. </p>
<p>When composing a formula, you can also reference a cell by clicking on the cell rather than typing out the cell address. This is particularly useful for multicell ranges (&quot;D4:G23&quot; for example). </p>
<h2>Editing a Formula</h2>
<p>Once you have entered a formula and you want to edit it, there are two options: You can select the cell containing the formula and then click in the Formula Bar at the top of the worksheet as follows:</p>
<p><img src="formula-edit-1.png" width="470" height="303" alt=""/></p>
<p>or you can double click on the cell containing the formula and edit it directly in the cell:</p>
<p><img src="formula-edit-2.png" width="401" height="129" alt=""/></p>
<h2>Functions</h2>
<p>One of the most powerful features of Excel is built-in functions. A function typically takes one or more arguments as input and returns a value. Functions are extremenly useful in formulas. For example, you can use trig functions:</p>
<blockquote>
  <p>sin(a)<br />
cos(a)    <br />
    tan(a)<br />
    etc.</p>
</blockquote>
<p>where a = an angle in radians. There are also many functions that operate on a range as input:</p>
<blockquote>
  <p>Sum(r)<br />
    Average(r)<br />
    Min(r)<br />
    Max(r)<br />
    etc.
  </p>
</blockquote>
<p>where r = a range of cells. For example, this formula computes the sum of a list of values:</p>
<p><img src="sum.png" width="334" height="413" alt=""/></p>
<p>A complete set of the available functions can be found in the Excel Help.</p>
<h2>Copying Formulas</h2>
<p>After entering a formula, it is often necessary to copy that formula to other cells. For example, the following spreadsheet is designed to compute the volume and weight of a set of cylinders defined by a radius and a height. The volume can be computed from the radius and height using the following formula:</p>
<p><img src="copy-formula-1.png" width="383" height="353" alt=""/></p>
<p>After entering the formula in cell D7, we wish to copy the formula to cells D8:D22. This can be accomplished by selecting cell D7 after the formula has been entered and copying (Ctrl-C) and pasting (Ctrl-V) the formula to D8:D22 using the clipboard. Another method is to select the cell as follows:</p>
<p><img src="copy-formula-2.png" width="386" height="131" alt=""/></p>
<p>and then drag the green square in the lower right corner of the cell down to the end of the list, or simply double-click on the green square. After doing so, the formula is automatically copied to the end of the list:</p>
<p><img src="copy-formula-3.png" width="398" height="384" alt=""/></p>
<h2>Relative vs. Absolute References</h2>
<p>When copying formulas, we need to be careful how were reference other cells in our formulas. For example, to calculate the weight of our cylinders, we take the volume of the cylinder and multiply by the unit wt of the cylinder material as follows:</p>
<p><img src="abs-rel-1.png" width="412" height="224" alt=""/></p>
<p>After copying the formula to the bottom of the table in column E, we notice that the weights are not properly computed:</p>
<p><img src="abs-rel-2.png" width="394" height="497" alt=""/></p>
<p>The reason for this error can be seen by revealing the formulas. This is accomplished by pressing Ctrl-~ on the keyboard (the &quot;~&quot; symbol is called the &quot;tilde&quot; and is on the upper left corner of your keyboard).</p>
<p><img src="abs-rel-3.png" width="507" height="251" alt=""/></p>
<p>Note that when you copy a formula, the cell references are updated with each subsequent cell the formula is copied to. Note that &quot;=B4*D7&quot; is changed to &quot;=B5*D8&quot; in the next cell down. This happens because whenever you reference a cell in a formula, that reference is interpreted to be <strong>relative</strong> to the cell containing the formula. In other words, when we type &quot;D7&quot; in a formula in cell E7, what we are really referencing is &quot;stay on the same row, but go one column to the left&quot;. Therefore, when the formula is copied, it correctly references the proper volume value one cell to the left. However, our error occurs because of a relative reference to the unit wt. value. A reference to B4 from cell E7 literally means &quot;three rows up and three columns to the left&quot;. But in this case, we don't want a relative reference. When we copy the formula, we want to ALWAYS reference cell B4. We can accomplish this by changing the B4 reference to make it <strong>absolute</strong> as follows:</p>
<p><img src="abs-rel-4.png" width="386" height="227" alt=""/></p>
<p>Note the &quot;$&quot; symbols. You make an absolute reference by directly typing the values or by typing B4 and then pressing the <strong>F4</strong> key (Command-T on a Mac). Now after copying the formula down, we get correct answers:</p>
<p><img src="abs-rel-5.png" width="393" height="486" alt=""/></p>
<p>And the formulas look like this:</p>
<p><img src="abs-rel-6.png" width="455" height="208" alt=""/></p>
<p>Sometimes it is useful to use a mixed reference. Here is a summary of the ways in which you can reference another cell.</p>
<blockquote>
  <table width="380" border="0">
    <tr>
      <td width="81">D4</td>
      <td width="289">Row and column are both relative</td>
    </tr>
    <tr>
      <td>$D4</td>
      <td>Row is relative and column is absolute</td>
    </tr>
    <tr>
      <td>D$4</td>
      <td>Row is absolute and column is relative</td>
    </tr>
    <tr>
      <td>$D$4</td>
      <td>Row and column are both absolute</td>
    </tr>
  </table>
</blockquote>
<p>For the example shown above, we could have gotten away with a mixed reference (&quot;B$4&quot;) because we copied the formulas within a single column, but it works fine with a complete absolute reference (&quot;$B$4&quot;). To do a mixed reference, you can either directly type the &quot;$&quot; symbols or you can repeatedly press the F4 key to get the combination you are seeking.</p>
<h2>Sample Workbook</h2>
<p>The workbook used in the examples shown above can be downloaded here:</p>
<p><a href="cylinders.xlsx">cylinders.xlsx</a></p>

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
      <td> <strong>The Basics -</strong> Run through some basic ways to input formulas into cells.</td>
      <td align="center" valign="top">Easy</td>
      <td align="center" valign="top"><a href="the_basics.xlsm">the_basics.xlsm</a></td>
      <td align="center" valign="top"><a href="the_basics_key.xlsm">the_basics_key.xlsm</a></td>
    </tr>
    <tr>
      <td><strong>Buoyancy- </strong>Calculate the buoyant force on different sized objects using formulas and cell references.</td>
      <td align="center" valign="top">Medium</td>
      <td align="center" valign="top"><a href="buoyancy.xlsm">buoyancy.xlsm</a></td>
      <td align="center" valign="top"><a href="buoyancy_key.xlsm">buoyancy_key.xlsm</a></td>
    </tr>
    <tr>
      <td><strong>Name Counter - </strong> Use a formula to count the number of names found inside a table range.</td>
      <td align="center" valign="top">Medium</td>
      <td align="center" valign="top"><a href="name_counter.xlsx">name_counter.xlsx</a></td>
      <td align="center" valign="top"><a href="name_counter_key.xlsx">name_counter_key.xlsx</a></td>
    </tr>
    <tr>
      <td><strong>Trigonometry - </strong>Use trigonometric functions inside of formulas to find the missing angles and/or sides of a few triangles.</td>
      <td align="center" valign="top">Medium</td>
      <td align="center" valign="top"><a href="trigonometry.xlsm">trigonometry.xlsm</a></td>
      <td align="center" valign="top"><a href="trigonometry_key.xlsm">trigonometry_key.xlsm</a></td>
    </tr>
    <tr>
      <td><strong>Employee Database -</strong> Use formulas to conduct statistics on an employee database.</td>
      <td align="center" valign="top">Hard</td>
      <td align="center" valign="top"><a href="employee_database.xlsm">employee_database.xlsm</a></td>
      <td align="center" valign="top"><a href="employee_database_key.xlsm">employee_database_key.xlsm</a></td>
    </tr>
  </tbody>
</table>
<br />

<?php 
require "../footer.php";
?>
</body>
</html>
