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

<h1>Naming Cells</h1>
<p>When writing  formulas in Excel, the formulas will often reference several other cells in the worksheet, or even cells on another worksheet. If the formula is complex, it can be difficult keep track of all of the references and determine if the formula is properly written. In such cases, we can make our formulas much simpler and more intuitive by using named cells. To illustrate this, consider the following workbook:</p>
<p><img src="start.png" width="636" height="534" alt=""/></p>
<p>The objective of the spreadsheet is to perform a set of calculations associated with a projectile fired from the coordinate origin (0,0) at an angle (&alpha;) at an initial velocity (v). A set of equations is then used to compute the range (r), the max height (h), and the total time in the air. The bottom table is used to compute a set of x,y coordinates defining the projective path. At the top of the sheet in cells C4:C7, a set of inputs are defined. The range, height, and total time are computed in cells C9:C11. Using the standard cell reference notation, the formulas in cells C9:C11 would appear as follows:</p>
<blockquote>
  <table width="273" border="0">
    <tr>
      <td>range</td>
      <td>=(C6^2*SIN(2*C5))/C7</td>
    </tr>
    <tr>
      <td>height</td>
      <td>=(C6^2*(SIN(C5))^2)/(2*C7)</td>
    </tr>
    <tr>
      <td>time</td>
      <td>=(2*C6*SIN(C5))/C7</td>
    </tr>
  </table>
</blockquote>
<p>These formulas can be compared to the equations shown in the above figure. To make our formulas easier to enter, read, and edit, we will now create a set of cells names and rewrite the formulas. Each cell has a default name based on the column-row combination (C6, D15, etc.). However, we can create an additional name (or alias) for a cell by selecting the cell and typing in a new name in the name box in the upper left corner just above the cells and below the menu. For example, to create a new name for the angle, we select cell C5 and then click in the name box and type in a new name (&quot;alpha&quot;) and then hit Return as follows:</p>
<p><img src="rename-alpha.png" width="402" height="316" alt=""/></p>
<p>We then repeat the same process for cells C5 (v) and C7 (g). Now, we can rewrite the formulas for the range, height, and time using the cell names. After doing so, they look like this:</p>
<blockquote>
  <table width="273" border="0">
    <tr>
      <td>range</td>
      <td>=(v^2*SIN(2*alpha))/g</td>
    </tr>
    <tr>
      <td>height</td>
      <td>=(v^2*(SIN(alpha))^2)/(2*g)</td>
    </tr>
    <tr>
      <td>time</td>
      <td>=(2*v*SIN(alpha))/g</td>
    </tr>
  </table>
</blockquote>
<p>Compare these formulas to the first set shown above and note how they more closely resemble the native set of equations shown in the first diagram. The names make the formulas easier to enter and easier to understand. This is especially true of longer formulas and cases where cells on different sheets are referenced.</p>
<h2>Relative vs. Absolute</h2>
<p>It is important to note that using a cell name in a formula represents an <strong>absolute</strong> reference. It is not possible to refer to a cell via a name as a relative reference unless you are using <a href="../excel-arrayform/">array formulas</a>. This is also an advantage to named cells because you can simply type the cell name without worrying about the &quot;$&quot; symbols associated with an absolute reference.</p>
<h2>Ranges</h2>
<p>Names can not only be applied to individual cells, but also to ranges of cells. For example, you can select the range B14:D10 and call it &quot;xytable&quot; or something. This is especially useful when writing formulas using the <a href="../excel-vlookup/">VLOOKUP</a> function where you need to make an absolute reference to a range of cells for the vlookup_table argument.</p>
<h2>Naming Rules</h2>
<p>Cell names can be any combination of letters and numbers, but you cannot use a name that would be otherwise interpreted as a native cell address (&quot;B4&quot;, &quot;C20&quot;, etc.). You also cannot put spaces or special characters (&quot;#&quot;, &quot;%&quot;, etc.) in the names. Single letter names (&quot;a&quot;, &quot;b&quot;, etc.) are allowed, except for &quot;c&quot; and &quot;r&quot; which are not allowed.</p>
<p>You can apply multiple names to the same cell or range. Any of the names can be used to reference the cell or range in a formula.</p>
<h2>Multiple Names</h2>
<p>You can assign as many names as you like to the same cell or range of cells. Any of the names can then be used to reference the cell or range in your formulas. Think of it as a person answering to multiple nicknames (<a href="https://www.youtube.com/watch?v=zhiqkq84pLA">"Champ", "Slugger", "Cowboy", "Buckaroo"</a>). </p>
<h2>Name Manager</h2>
<p>In addition to using the name box as described above, you can also name a cell using the <strong>Name Manager</strong> button located in the Fomulas ribbon:</p>
<p><img src="namebutton.png" width="391" height="187" alt=""/></p>
<p>This brings up the following dialog:</p>
<p><img src="namemanager.png" width="547" height="351" alt=""/></p>
<p>This dialog can be used to edit and/or delete names, or to create new names.</p>
<h2>Deleting a Name</h2>
<p>One needs to be careful when deleting a name associated with a cell or range. You CANNOT select the cell or range and then delete the name from the name box in the upper left corner of the worksheet. This may appear to delete the name but it does not. The only way to delete a name is using the <strong>Name Manager</strong>.</p>
<h2>Sample Workbook</h2>
<p>The workbook used in the examples shown above can be downloaded here:</p>
<p><a href="projectile.xlsx">projectile.xlsx</a></p>

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
      <td> <strong>Names -</strong> Use cell naming to find the volume and weight of different objects.</td>
      <td align="center" valign="top">Easy</td>
      <td align="center" valign="top"><a href="names.xlsm">names.xlsm</a></td>
      <td align="center" valign="top"><a href="names_key.xlsm">names_key.xlsm</a></td>
    </tr>
    <tr>
      <td><strong>Moment Arm - </strong>Use names as inputs to a formula to calculate the moment arm on a lever at different lengths.</td>
      <td align="center" valign="top">Medium</td>
      <td align="center" valign="top"><a href="moment_arm.xlsm">moment_arm.xlsm</a></td>
      <td align="center" valign="top"><a href="moment_arm_key.xlsm">moment_arm_key.xlsm</a></td>
    </tr>
    <tr>
      <td><strong>Bernoulli Problem -</strong> Use names inside of formulas to determine the solution to the bernoulli equation. </td>
      <td align="center" valign="top">Hard</td>
      <td align="center" valign="top"><a href="bernoulli_problem.xlsm">bernoulli_problem.xlsm</a></td>
      <td align="center" valign="top"><a href="bernoulli_problem_key.xlsm">bernoulli_problem_key.xlsm</a></td>
    </tr>
  </tbody>
</table>
<br />

<?php 
require "../footer.php";
?>
</body>
</html>
