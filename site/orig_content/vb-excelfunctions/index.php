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
<h1>Calling Excel Functions from VB Code</h1>
<p>One of the nice things about writing VB code inside Excel is that you can 
  combine all of the power and flexibility of Visual Basic with the many tools and 
  options in Excel. One of the best examples of this is that you can take 
  advantage of all of the standard Excel worksheet functions inside your VB code.&nbsp;Calling an Excel worksheet function is simple. The Excel functions are 
  available as methods within the <b>WorksheetFunction</b> object. You 
  simply invoke the method and pass the arguments that the function requires 
  (typically a range).&nbsp; </p>
<p>For example, if we were writing a simple formula to put in a cell to find the 
  minimum value in a range of cells, we would write the following:</p>

<pre><code class="language-vb">=Min(B4:F30)
</code></pre>
  

<p>The following code uses the same <b>Min</b> function, but invokes the 
  function using VB code.&nbsp; The min value is stored in a variable called <b> minval</b>:</p>

<pre><code class="language-vb">Dim minval As Double
minval = Application.WorksheetFunction.Min(Range("B4:F30"))
</code></pre>  
  
<p>Notice the difference in how the range is specified.&nbsp; In the VB code, 
  the range is specified as a range object.</p>
<p>The Application. portion is actually optional and can be omitted in most 
  cases.&nbsp; Thus, the following code achieves the same thing:</p>

<pre><code class="language-vb">Dim minval As Double
minval = WorksheetFunction.Min(Range("B4:F30"))
</code></pre>  
  
<p>Here are some more examples:</p>

<pre><code class="language-vb">Range("e5") = WorksheetFunction.sum(Range("b5:b29"))

'This is useful since VB does not have an inverse sin function
Dim x As Double
x = WorksheetFunction.Asin(0.223)

Dim i As Integer
i = 5
Range("H4") = WorksheetFunction.Fact(i)
</code></pre>

<br />

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
      <td> <strong>Harmonic Mean -</strong> Use an Excel function within a custom function to calculate the harmonic mean from the tabulated data.</td>
      <td align="center" valign="top">Easy</td>
      <td align="center" valign="top"><a href="harmonic_mean.xlsm">harmonic_mean.xlsm</a></td>
      <td align="center" valign="top"><a href="harmonic_mean_key.xlsm">harmonic_mean_key.xlsm</a></td>
    </tr>
    <tr>
      <td><strong>Law of Cosines - </strong>Calculate the Law of Cosines using an Excel function for Cosine within your sub.</td>
      <td align="center" valign="top">Medium</td>
      <td align="center" valign="top"><a href="law_of_cosines.xlsm">law_of_cosines.xlsm</a></td>
      <td align="center" valign="top"><a href="law_of_cosines_key.xlsm">law_of_cosines_key.xlsm</a></td>
    </tr>
    <tr>
      <td><strong>Bill Payment -</strong> Use the APR Excel function within a custom function to calculate the number of months required to pay off a credit card bill.</td>
      <td align="center" valign="top">Hard</td>
      <td align="center" valign="top"><a href="bill_payments.xlsm">bill_payments.xlsm</a></td>
      <td align="center" valign="top"><a href="bill_payments_key.xlsm">bill_payments_key.xlsm</a></td>
    </tr>
  </tbody>
</table>
<br />

<?php 
require "../footer.php";
?>
</body>
</html>
