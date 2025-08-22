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

<h1> Using the IF Function</h1>

<p>There are many instances when using Excel where we need to write a formula that  produces one of two different results, depending on some condition. For example, consider the following version of our cylinder analysis worksheet:</p>
<p><img src="workbook-1.png" width="482" height="490" alt=""/></p>
<p>Suppose we wish to add a &quot;Class&quot; item in column F indicating whether the cylinders are standard or overweight. This can be done by entering the following formula in cell F7 and copying it down to the end of the table:</p>
<blockquote>
  <p>=IF(E7&lt;35000, &quot;Standard&quot;, &quot;Overweight&quot;)</p>
</blockquote>
<p>Resulting in the following:</p>
<p><img src="workbook-2.png" width="473" height="486" alt=""/></p>
<h2>Syntax</h2>
<p>The syntax of the IF function is as follows:</p>
<blockquote>
  <p>=	IF(logical_test, value_if_true, value_if_false)</p>
</blockquote>
<p>The <b>logical_test</b> argument needs to be a conditional expression that returns a TRUE or FALSE value. Conditional expression are typically formed with one of the following operators:</p>
<table width="357" border="1">
  <tr>
    <td width="76" align="center"><strong>Operator</strong></td>
    <td width="91" align="center"><strong>Example</strong></td>
    <td width="176"><strong>Description</strong></td>
  </tr>
  <tr>
    <td align="center">=</td>
    <td align="center">A4=0</td>
    <td>Equal</td>
  </tr>
  <tr>
    <td align="center">&lt;&gt;</td>
    <td align="center">A4&lt;&gt;B5</td>
    <td>Not equal to</td>
  </tr>
  <tr>
    <td align="center">&gt;</td>
    <td align="center">D7&gt;3</td>
    <td>Greater than</td>
  </tr>
  <tr>
    <td align="center">&gt;=</td>
    <td align="center">D4&gt;=0</td>
    <td>Greater than or equal to</td>
  </tr>
  <tr>
    <td align="center">&lt;</td>
    <td align="center">G3&lt;(G4-7)</td>
    <td>Less than</td>
  </tr>
  <tr>
    <td align="center">&lt;=</td>
    <td align="center">0&lt;=F12</td>
    <td>Less than or equal to</td>
  </tr>
</table>
<p>If the conditional expression evaluates to true, the <strong>value_if_true</strong> argument is used. Otherwise, the <strong>value_if_false</strong> argument is used. These arguments can any type of expression, including constants, cells references, or formulas. Here are some additional example formulas that use the IF function:</p>
<blockquote>
  <p>=IF(A4&lt;&gt;0, 1/A4, &quot;Error - Divide by Zero!&quot;)</p>
  <p>=IF(B4&lt;=$D$2, -2.3*G4/4, -3.9*G4/4+6)</p>
  <p>=IF(units=&quot;Metric&quot;, &quot;[m/sec], &quot;[ft/sec]&quot;)</p>
</blockquote>
<h2>Compound Conditions</h2>
<p>Sometimes we need to utilize compound conditional expressions with the IF function. But we need to be very careful when doing so. For example, suppose we want to represent the following mathematical expression:</p>
<blockquote>
  <p>0 &le; x &le; 5</p>
</blockquote>
<p>in an Excel formula and &quot;x&quot; is stored in cell B5. It would be tempting to use the following conditional expression:</p>
<blockquote>
  <p>0&lt;=B5&lt;=5</p>
</blockquote>
<p>for the first argument in the IF function. However, this creates a useless and incorrect expression that will always return TRUE, regardless of the contents of cell B5. This is because a compound expression like this is evaluated one operator at a time from left to right. In other words, the first part of the expression:</p>
<blockquote>
  <p><strong>0&lt;=B5</strong>&lt;=5</p>
</blockquote>
<p>will be evaluated first. The result of this evaluation will be True or False, depending on whether or not B5 is greater than or equal to zero. This result is then compared against the rest of the expression. In computational terms, True and False evaluate to 1 or 0, respectively. Thus, if B5 is greater than or equal to zero, the expression simplifies to:</p>
<blockquote>
  <p>1&lt;=5</p>
</blockquote>
<p>otherwise (B5&gt;0), it simplifies to:</p>
<blockquote>
  <p>0&lt;=5</p>
</blockquote>
<p>Both of these statements will then evaluate to True, regardless of the value of B5. In other words, the original expresssion is equivalent to:</p>
<blockquote>
  <p>(0&lt;=B5)&lt;=5</p>
</blockquote>
<p>which is fundamentally different from the mathematical expression we are trying to emulate. To solve this problem correctly, we need to use the <strong>AND</strong> function as follows:</p>
<blockquote>
  <p>AND(0&lt;=B5, B5&lt;=5)</p>
</blockquote>
<p>This function returns True if both statements are true. Otherwise it returns False. Likewise, there is an <strong>OR</strong> function that returns True if either or both of the two expressions evaluate to True.</p>
<h2>Nested IF Functions</h2>
<p>It is possible to nest multiple instances of the IF function. For example:</p>
<blockquote>
  <p>=IF(A4&gt;=18,&quot;Adult&quot;,IF(A4&gt;12,&quot;Teen&quot;,&quot;Child&quot;))</p>
</blockquote>
<p>The second IF function is only evaluated in the first condition is False. There are three possible outcomes in this case: &quot;Adult&quot;, &quot;Teen&quot;, and &quot;Child&quot;.</p>
<h2>Sample Workbook</h2>
<p>The workbook used in the first example shown above can be downloaded here:</p>
<p><a href="cylinders3.xlsx">cylinders3.xlsx</a></p>

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
      <td> <strong>Reynolds and Froude -</strong> Calculate the Reynolds' or Froude's number by inputing an IF equation into the appropriate cell. </td>
      <td align="center" valign="top">Easy</td>
      <td align="center" valign="top"><a href="reynolds_and_froude.xlsx">reynolds_and_froudes.xlsx</a></td>
      <td align="center" valign="top"><a href="reynolds_and_froude_key.xlsx.xlsm">reynolds_and_froudes_key.xlsx</a></td>
    </tr>
    <tr>
      <td><strong>Concrete Price Estimator - </strong>Use an IF equation to determine the varying prices of different concrete projects.</td>
      <td align="center" valign="top">Medium</td>
      <td align="center" valign="top"><a href="concrete_price_estimator.xlsx">concrete_price_estimator.xlsx</a></td>
      <td align="center" valign="top"><a href="concrete_price_estimator_key.xlsx">concrete_price_estimator_key.xlsx</a></td>
    </tr>
    <tr>
      <td><strong>Parking - </strong>Determine which types and how many vehicles you can park along side a given curb.</td>
      <td align="center" valign="top">Medium</td>
      <td align="center" valign="top"><a href="parking.xlsx">parking.xlsx</a></td>
      <td align="center" valign="top"><a href="parking_key.xlsx">parking_key.xlsx</a></td>
    </tr>
    <tr>
      <td><strong>Years Left of School -</strong> Use IF statements and user inputs/selections to determine how many years that the user has left to finish school.</td>
      <td align="center" valign="top">Hard</td>
      <td align="center" valign="top"><a href="years_left_of_school.xlsm">years_left_of_school.xlsx</a></td>
      <td align="center" valign="top"><a href="years_left_of_school_key.xlsm">years_left_of_school_key.xlsx</a></td>
    </tr>
  </tbody>
</table>
<br />

<?php 
require "../footer.php";
?>
</body>
</html>
