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

<h1>Operators and Precedence - Writing Complex Formulas</h1>
<p>When developing engineering applications with Excel, we often take complex equations and convert them to Excel formulas. In so doing, it is easy to compose the formula in such a way that leads to incorrect results. For example, consider the following spreadsheet:</p>
<p>
  <!--td {border: 1px solid #ccc;}br {mso-data-placement:same-cell;}-->
  <!--td {border: 1px solid #ccc;}br {mso-data-placement:same-cell;}--><img src="quadsheet_start.png" width="404" height="417" alt=""/></p>
<p>The objective of the spreadsheet is to solve for the roots of a quadratic equation of the form:</p>
<blockquote>
  <p>ax<sup>2</sup> + bx + c = 0</p>
</blockquote>
<p>using the two equations shown above. Let's focus on the euqation for root 1 which would be entered in cell <strong>D5</strong>. How should one go about transforming the native equation into a formula that is properly interpreted by Excel? Consider the following potential solution:</p>
<blockquote>
  <p>=-B5+B5^2-4*A5*C5^0.5/2*A5</p>
</blockquote>
<p>This results in a solution of:</p>
<p><img src="numerror.png" width="404" height="190" alt=""/></p>
<p>The most obvious error is that we need to put parentheses around the discriminant (b2 - 4ac) part before taking the square root ( we are taking the square root by raising to a power of 0.5). Otherwise we are only taking the square root of c only. After making this correction:</p>
<blockquote>
  <p>=-B5+(B5^2-4*A5*C5)^0.5/2*A5</p>
</blockquote>
<p>which gives us an answer of 6.58, which is still wrong. The correct answer is 1.65. In order to get the correct answer, we need to put parentheses around both the numerator and denominator as follows:</p>
<blockquote>
  <p>=(-B5+(B5^2-4*A5*C5)^0.5)/(2*A5)</p>
</blockquote>
<p>After doing so, we finally get the correct set of answers:</p>
<p><img src="solution.png" width="378" height="268" alt=""/></p>
<p>So how do we determine when and where to use parentheses? Are they always required? Here is another example. Take the following formula that uses cell names:</p>
<blockquote>
  <p>=x+y*z^p/2*x</p>
</blockquote>
<p>Which of the following  corresponds to the equation defined by this formula?</p>
<blockquote>
  <table width="200" border="0">
    <tr>
      <td>a.</td>
      <td><img src="equation1.png" width="103" height="56" alt=""/></td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>b.</td>
      <td><img src="equation2.png" width="86" height="62" alt=""/></td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>c.</td>
      <td><img src="equation3.png" width="87" height="61" alt=""/></td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>d.</td>
      <td><img src="equation4.png" width="104" height="60" alt=""/></td>
    </tr>
  </table>
</blockquote>
<p>The correct answer is (d). But how do we know this? One can always very explicitly define the order of operations by using parentheses. But we can better understand how and when to use parentheses by learning the precedence and associativity rules applied to Excel formulas. The following table indicates the order in which the operators are evaluated when a formular is parsed.</p>
<blockquote>
  <table width="293" border="0">
    <tr>
      <td width="20">1.</td>
      <td width="58" align="center">^</td>
      <td width="201">Power (exponent)</td>
    </tr>
    <tr>
      <td>2.</td>
      <td align="center">* /</td>
      <td>Multiplication and division</td>
    </tr>
    <tr>
      <td>3.</td>
      <td align="center">+ -</td>
      <td>Addition and subtraction</td>
    </tr>
  </table>
</blockquote>
<p>Within levels 2 and 3, operations are carried out from left to right. So, let's reexamine our formula from above:</p>
<blockquote>
  <p>=x+y*z^p/2*x</p>
</blockquote>
<p>The first step would be to evaluate the power operator (^):</p>
<blockquote>
  <p>=x+y*(z^p)/2*x</p>
</blockquote>
<p>Next, the multiplication (*) and division (/) operators would be evaluated from left to right.</p>
<blockquote>
  <p>=x+(y*(z^p))/2*x</p>

  <p>=x+((y*(z^p))/2)*x</p>
  <p>=x+(((y*(z^p))/2)*x)</p>
</blockquote>
<p>Finally, the addition (+) operator would be evaluated last.</p>
<blockquote>
  <p>	=x+(((y*(z^p))/2)*x)</p>
</blockquote>
<p>Of course, these parentheses are not all required. Once we understand the precedence rules, we can begin to build concise, but correct formulas. For example, this is how you would write a formula for each of the equations shown above:</p>
<blockquote>
  <table width="293" border="0">
    <tr>
      <td width="12">a.</td>
      <td width="121"><img src="equation1.png" width="103" height="56" alt=""/></td>
      <td width="146">=x+y*z^(p/(2*x))</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>b.</td>
      <td><img src="equation2.png" width="86" height="62" alt=""/></td>
      <td>=(x+y*z^p)/(2*x)</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>c.</td>
      <td><img src="equation3.png" width="87" height="61" alt=""/></td>
      <td>=x+(y*z^p)/(2*x)</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>d.</td>
      <td><img src="equation4.png" width="104" height="60" alt=""/></td>
      <td>=x+y*z^p/2*x</td>
    </tr>
  </table>
</blockquote>
<p>The worksheet associated with the exercises shown above can be downloaded here:</p>
<blockquote>
  <p><a href="quadequation.xlsm">quadequation.xlsm</a></p>
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
      <td> <strong>Headloss -</strong> Use the order of operations to calculate the fluid headloss in a pipe. </td>
      <td align="center" valign="top">Easy</td>
      <td align="center" valign="top"><a href="headloss.xlsm">headloss.xlsm</a></td>
      <td align="center" valign="top"><a href="headloss_key.xlsm">headloss_key.xlsm</a></td>
    </tr>
    <tr>
      <td><strong>Operators - </strong>Calculate the equations using formulas and taking into consideration the appropriate order of operations.</td>
      <td align="center" valign="top">Medium</td>
      <td align="center" valign="top"><a href="operators.xlsm">operators.xlsm</a></td>
      <td align="center" valign="top"><a href="operators_key.xlsm">operators_key.xlsm</a></td>
    </tr>
    <tr>
      <td><strong>Friction Factor -</strong> Use the correct order of operations to calculate the friction factor.</td>
      <td align="center" valign="top">Hard</td>
      <td align="center" valign="top"><a href="friction_factor.xlsm">friction_factor.xlsm</a></td>
      <td align="center" valign="top"><a href="friction_factor_key.xlsm">friction_factor_key.xlsm</a></td>
    </tr>
  </tbody>
</table>
<br />

<?php 
require "../footer.php";
?>
</body>
</html>
