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

<h1>Data Validation</h1>
<p>In order for our formulas to work properly, it is often helpful to set limitations on what can be entered into some of our input cells. This can be easily accomplished in Excel using the data validation tools. To illustrate how this works, we will use the Cylinder Analysis example features in the <a href="../excel-vlookup/index.php">VLOOKUP</a> chapter. In that example, we have a table of unit weights for a set of selected materials and in the lower table we have a column of Materiasl (column E) where the user will enter a material for each of the cylinders. </p>
<p><img src="../excel-vlookup/start.png" width="473" height="609" alt=""/></p>
<p>In order for things to work properly, each of the entries in the Material column needs to match one of the entries in the first column of the unit weight vs. material table at the top of the sheet. We can ensure that this happen by applying data validation to the Material column in the lower table. First we need to select the cells in the Materials column (E12:E28). Then we select the <strong>Data Valdation</strong> button in the <strong>Data</strong> tab. This brings up the Data Validation dialog where we enter the following:</p>
<p><img src="validation-1.png" width="891" height="610" alt=""/></p>
<p>Note that we have selected the <strong>List</strong> item in the <strong>Allow</strong> options. This means that we will allow the user to select an item from a list. Then we enter the address of the list (i.e., the first column in the unit weight table) in the Source field. This can be done by directly typing the formula as shown or simply by putting the cursor in the field and then selecting the range of cells. After clicking OK, whenever the user selects one of the cells in the column to enter a value, a pop-up menu is presented:</p>
<p><img src="validation-2.png" width="263" height="147" alt=""/></p>
<p>If one of the items in the list is not selected, an error message is given. We can now populate the entire list with materials.</p>
<p>It should be noted that Data Validation can be used be used for all types of checks on the input. The Allow options are as follows:</p>
<p><img src="validation-3.png" width="408" height="325" alt=""/></p>
<p>For example, if the Decimal option is selected, the following options are presented:</p>
<p><img src="validation-4.png" width="408" height="325" alt=""/></p>
<p>The <strong>Data</strong> option can be used to select &quot;between&quot;, &quot;greater than&quot;, &quot;less than&quot;, etc. Thus, we can carefully control what values are allowed into each of our input cells, thereby minimizing the chance of errors.</p>

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
      <td> <strong>Inventory -</strong> Validate columns of a table used to keep track of inventory of various construction items.</td>
      <td align="center" valign="top">Easy</td>
      <td align="center" valign="top"><a href="inventory.xlsx">inventory.xlsx</a></td>
      <td align="center" valign="top"><a href="inventory_key.xlsx">inventory_key.xlsx</a></td>
    </tr>
    <tr>
      <td><strong>Dog Show - </strong>Use data validation to control the user inputs on a spreadsheet used for scoring at a dog show.</td>
      <td align="center" valign="top">Medium</td>
      <td align="center" valign="top"><a href="dog_show.xlsx">dog_show.xlsx</a></td>
      <td align="center" valign="top"><a href="dog_show_key.xlsx">dog_show_key.xlsx</a></td>
    </tr>
  </tbody>
</table>
<br />

<?php 
require "../footer.php";
?>
</body>
</html>
