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

<h1> Array Formulas</h1>

<p>In Excel terminology, an array is a group of cells in a row (horizonal array), column (vertical array), or multiple rows and columns (2D array). In this chapter we discuss a special type of Excel formula called an &quot;array formula&quot; that operates on arrays. Array formulas can be a powerful tool for condensing a complex set of calculations in a simple, concise formula.</p>
<p>To illustrate the power of array formulas, we will use the following workbook as an example:</p>
<p><img src="sheet-start.png" width="555" height="571" alt=""/></p>
<p>The workbook is designed to perform calculations associated with the sale of a set of products. Each column represents a sale. The product ID is in column A, the number units sold in the sale is in column B, and the wholesale and retail prices are in columns C and D. We will calculate the total sale amount in terms of wholesale and retail in columns E and F and then perform some summary computations in the section at the bottom.</p>
<p>When dealing with array formulas, it is sometimes useful to use named ranges. This is not a requirement but it makes the formulas much easier to understand. For the purpose of this example, we will name the following set of ranges:</p>
<table width="294" border="1">
  <tr>
    <th width="59" align="center" scope="col">Range</th>
    <th width="52" align="center" scope="col">Column</th>
    <th width="52" align="center" scope="col">Name</th>
  </tr>
  <tr>
    <td align="center">B6:B17</td>
    <td align="center">Units</td>
    <td align="center">units</td>
  </tr>
  <tr>
    <td align="center">C6:C17</td>
    <td align="center">Price|Wholesale</td>
    <td align="center">pw</td>
  </tr>
  <tr>
    <td align="center">D6:D17</td>
    <td align="center">Price|Retail</td>
    <td align="center">pr</td>
  </tr>
  <tr>
    <td align="center">E6:E17</td>
    <td align="center">Totals|Wholesale</td>
    <td align="center">tw</td>
  </tr>
  <tr>
    <td align="center">F6:F17</td>
    <td align="center">Totals|Retail</td>
    <td align="center">tr</td>
  </tr>
</table>
<p>These names will be used in each of the formula examples shown below.</p>
<h2>Simple Array Calculations</h2>
<p>To begin, we will use an array formula to calcute the wholesale total in column E. Normally, we will do this by entering the following formula cell E6:</p>
<blockquote>
  <p>=B6*C6</p>
</blockquote>
<p>and then copying the formula to the rest of the cells in column E. To do this with an array formula, we first select the the entire Total|Wholesale column (E6:E17) and then type the following in the formula bar:</p>
<blockquote>
  <p>=units*pw</p>
</blockquote>
<p>and then rather than simply hitting the Enter key, we hit <strong>Ctrl-Shift-Enter</strong>. We must always use this key sequence when entering an array formula. Otherwise we get either an error message or the wrong answer, depending on what is selected. After entering the array formula, the entire column is populated as follows:</p>
<p><img src="form-twcolumn.png" width="533" height="427" alt=""/></p>
<p>Note the resulting format of the formula:</p>
<blockquote>
  <p>={units*pw}</p>
</blockquote>
<p>The curly braces indicate that is is an array formula. If we wish to edit the formula, we can click on the formula bar and make changes, but we must always hit <strong>Ctrl-Shift-Enter</strong> when we are done.</p>
<p><em><strong>Note: </strong>If you are using excel through a Microsoft 365 subscription, you may not need to hit <strong>control-shift-enter </strong>to create an array formula. However, it is good to know how to so that if you are ever on a different version of excel</em> <em>that requires the <strong>control-shift-enter </strong>method, you will still be able to use array formulas.</em></p>
<p>An array formula is similar to peforming vector algebra. The formula essentially multiplies the units column by the wholesale price column. The result of multiplying these two vertical arrays is a vertical array of same dimension (12 items) where each item it equal to the number of units times the price for that particular entry (row in this case). The equation applies to the entire column in the <strong>Totals|Wholesale</strong> part of the table. One advantage of using an array formula in a case like this is that the resulting formula is simple and intuitive.</p>
<p>Next, we will do the same thing for the <strong>Total|Retail</strong> column. We will select the column and enter the following formula in the formula bar:</p>
<blockquote>
  <p>=units*pr</p>
</blockquote>
<p>and finish with <strong>Ctrl-Shift-Enter</strong>. At this point, the table is complete:</p>
<p><img src="form-trcolumn.png" width="542" height="433" alt=""/></p>
<h2>Using Arrays with Functions</h2>
<p>Now we will focus on the bottom section where we will perform some summary calculations based on our sales totals using functions. There are many Excel functions such as Sum() or Average() that take an array as input and return a single number as output. We will use these formulas in combination with array algebra to create some interesting results.</p>
<p>First of all, we calculate the wholesale total by calculating the sum of the number of units times the wholesale price. We could do that by calculating the sum of the Totats|Wholesale column, but with an array formula, we can peform the calculation without using this column. This is one of the advantages of array formulas is that we do not need any intermediate columns in order to some calculations involving multi-cell ranges. We will enter the following array formula in cell <strong>D20</strong>:</p>
<blockquote>
  <p>=SUM(units*pw)</p>
</blockquote>
<p>and we get value of $4,865.00. Similarly, we can use the formula:</p>
<blockquote>
  <p>=SUM(units*pr)</p>
</blockquote>
<p>to calculate the retail total (). Once again, after typing the formulas, we finish with <strong>Ctrl-Shift-Enter</strong> and the formulas are displayed with curly braces.</p>
<blockquote>
  <p>{=SUM(units*pw)}</p>
</blockquote>
<blockquote>
  <p>{=SUM(units*pr)}</p>
</blockquote>
<p>Next we will calculate the total profit as the sum of the retail totals minus the sum of the wholesale totals in cell <strong>D22</strong> using the following formula.</p>
<blockquote>
  <p>=SUM(tr-tw)</p>
</blockquote>
<p>Note that we use the wholesale  totals and retail totals columns in this calculation. But we could have performed the calculations directly from the units column and the prices columns as follows:</p>
<blockquote>
  <p>=SUM(units*pr-units*pw)</p>
</blockquote>
<p>In cell <strong>D23</strong>, we will calculate the maximum markup (difference between retail and wholesale prices) as follows:</p>
<blockquote>
  <p>=MAX(pr-pw)</p>
</blockquote>
<p>And the average markup as:</p>
<blockquote>
  <p>=AVERAGE(pr-pw)</p>
</blockquote>
<p>The maximum profit (difference between retail and wholesale totals) can be computed as:</p>
<blockquote>
  <p>=MAX(tr-tw)</p>
</blockquote>
<p>Finally, in cell <strong>D26</strong> we wish to compute the total high-price profit. The high-price profit is defined as the profit on items were the retail price is greater than $50. In order to calculate this correctly we need to compute a sum of only those items where the prices greater than $50. In order to do this we need to combine both the IF funciton and the SUM function as follows:</p>
<blockquote>
  <p>=SUM(IF(pr&gt;50,tr-tw,0))</p>
</blockquote>
<p>This example is a nice illustration of the power of array formulas. Doing this type of calculation within a normal formula would've been much more difficult than it was in this case.</p>
<p>At this point, our sales summary is complete:</p>
<p><img src="summary.png" width="258" height="148" alt=""/></p>
<h2>Logical Functions</h2>
<p>One caveat associated with using array formulas in Excel is that you have to be very careful when using logical functions such as OR() and AND(). For some reason, they don't behave the way you would normally expect, and can lead to logical errors in your formula results. However there is a way you can reformulate your array formulas to work around this limitation. This issue and the associated workarounds are explained in the following article:</p>
<blockquote>
  <p><a href="http://dailydoseofexcel.com/archives/2004/12/04/logical-operations-in-array-formulas/">http://dailydoseofexcel.com/archives/2004/12/04/logical-operations-in-array-formulas/</a></p>
</blockquote>
<p> The workarounds rely on the fact that the value of <strong>True</strong> is equal to the numerical value of <strong>1</strong> and a value of <strong>False</strong> is equal to the numerical value of <strong>0</strong>.</p>
<h2>Sample Workbook</h2>
<p>The workbook used in the examples shown above can be downloaded here:</p>
<p><a href="arrayform.xlsx">arrayform.xlsx</a></p>

<h2>Exercises</h2>
<p>You may wish to complete following exercises to  gain practice with and reinforce  the topics covered in this chapter:</p>
<table width="900" border="1">
  <tbody>
    <tr>
      <td width="312"><strong>Description</strong></td>
      <td width="84" align="center"><strong>Difficulty</strong></td>
      <td width="161" align="center"><strong>Start</strong></td>
      <td width="192" align="center"><strong>Solution</strong></td>
    </tr>
    <tr>
      <td> <strong>Financial Arrays -</strong> Conduct a simple financial analysis on a list of items using arrays. This worksheet will help explain/show the benefits of using arrays. </td>
      <td align="center" valign="top">Easy</td>
      <td align="center" valign="top"><a href="financial_arrays.xlsx">financial_arrays.xlsx</a></td>
      <td align="center" valign="top"><a href="financial_arrays_key.xlsx">financial_arrays_key.xlsx</a></td>
    </tr>
    <tr>
      <td><strong>New Car Expense Calculator - </strong>Calculate a range of expenses for buying a new car. </td>
      <td align="center" valign="top">Medium</td>
      <td align="center" valign="top"><a href="new_car_expense_calculator.xlsx">new_car_expense_calculator.xlsx</a></td>
      <td align="center" valign="top"><a href="new_car_expense_calculator_key.xlsx">new_car_expense_calculator_key.xlsx</a></td>
    </tr>
    <tr>
      <td><strong>Force Analysis -</strong> Use arrays to find different statistics of force from test data.</td>
      <td align="center" valign="top">Hard</td>
      <td align="center" valign="top"><a href="force_analysis.xlsx">force_analysis.xlsx</a></td>
      <td align="center" valign="top"><a href="force_analysis_key.xlsx">force_analysis_key.xlsx</a></td>
    </tr>
  </tbody>
</table>
<br />

<?php 
require "../footer.php";
?>
</body>
</html>
