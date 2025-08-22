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
<h1>Decisions and Conditions - Writing If Statements</h1>
<p>There are countless ocassions when writing code where we need to execute code only if certain conditions are met. This can be accomplished by writing <strong>If</strong> statements. </p>
<h2>Syntax</h2>
<p>The general syntax of an If statement is as follows:</p>

<pre><code class="language-vb">If condition_1 Then
    statement(s)
ElseIf condition_2 Then 
    statement(s)

...

ElseIf condition_n Then
    statement(s)
Else
    statement(s)
End If
</code></pre>

<p>The <strong>If ... Then</strong> and the <strong>End If</strong> parts are required. All of the other parts are optional and are only used when needed. For example, in some cases you don't need any of the Else options:</p>

<pre><code class="language-vb">If Range("B5") = "" Then
    MsgBox "Error: Cell B5 cannot be empty"
    Exit Sub
End If
</code></pre>


<p>Note that you can put as many statements as you want between the <strong>If</strong> and <strong>End If</strong> lines. Each of the statements is executed if the condition is true. You do not have to indent, but it is <strong>STRONGLY</strong> recommended as it makes your code much easier to follow.</p>
<p>In some cases you may wish to include an <strong>Else</strong> clause that is executed when the condition is false.</p>

<pre><code class="language-vb">If x = 0 Then
    MsgBox "Error: Cannot divide by zero"
Else
    y = 1 / x
End If</code></pre>

<p>In other cases, you may need to check several conditions, each of which is mutually exclusive.</p>

<pre><code class="language-vb">If yourteam = "BYU" Then
    MsgBox "You are cool"
ElseIf yourteam = "Utah" Then
    MsgBox "You are NOT cool"
ElseIf yourteam = "Utah State" Then
    MsgBox "What is an aggie?"
Else
    MsgBox "I do not care"
End If
</code></pre>

<p>Each of the conditions is checked in sequence starting at the top. Once a condition is found that evaluates to True, none of the remaining conditions are tested and the flow of control exits the IF statement and jumps to the code immediately following the End If statement.</p>

<p>For the cases with no <strong>ElseIf</strong> clauses and simple one-line results, you can put your entire statement on a single line:</p>

<pre><code class="language-vb">If yourteam = "BYU" Then MsgBox "You are cool"
</code></pre>

<p>or</p>

<pre><code class="language-vb">If x = 0 Then y = 0 Else y = 1 / x
</code></pre>

<h2>Conditional Expressions</h2>
<p>Every If statement requires at least one conditional expression. A conditional expression is an expression that returns either True or False when evaluated. Conditional expressions are generally formulated using a binary conditional operator. A binary operator takes two arguments, one on each side of the operator. Here is a list of the commonly used operators:</p>
<table width="341" border="0">
  <tr>
    <td width="188"><strong>Operator</strong></td>
    <td width="53" align="center"><strong>Symbol</strong></td>
    <td width="86" align="center"><strong>Example</strong></td>
  </tr>
  <tr>
    <td>Equal</td>
    <td align="center">=</td>
    <td align="center">a = b</td>
  </tr>
  <tr>
    <td>Not equal</td>
    <td align="center">&lt;&gt;</td>
    <td align="center">a &lt;&gt; b</td>
  </tr>
  <tr>
    <td>Less than</td>
    <td align="center">&lt;</td>
    <td align="center">x &lt; y</td>
  </tr>
  <tr>
    <td>Greater than</td>
    <td align="center">&gt;</td>
    <td align="center">p &gt; q</td>
  </tr>
  <tr>
    <td>Less than Or equal to</td>
    <td align="center">&lt;=</td>
    <td align="center">x &lt;= 5.5</td>
  </tr>
  <tr>
    <td>Greater than or equal to</td>
    <td align="center">&gt;=</td>
    <td align="center">y &gt;= p</td>
  </tr>
</table>
<p>Multiple conditional expressions can be combined with the <strong>And</strong> and <strong>Or</strong> operators. With the And operator, the combined expression is true if both conditions are true. With the Or operator, the combined expression is true if either of the two conditions is true. For example,</p>

<pre><code class="language-vb">If myteam = "BYU" And yourteam = "BYU" Then
    MsgBox "High five!"
ElseIf yourteam = "Utah" Or yourteam = "USU" Then
    MsgBox "Boo!"
Else
    MsgBox "Nice to meet you"
End If
</code></pre>

<p>You may wish to combine more than two conditional expressions. In this case, it helps to use parentheses.</p>

<pre><code class="language-vb">If (myteam = "BYU") And (yourteam = "Utah" Or yourteam = "USU") Then
    MsgBox "We are going to have a problem."
Else
    MsgBox "Nice to meet you"
End If
</code></pre>

<p>You can also use the <strong>Not</strong> operator to negate a statement. It is a unary operator and it negates the conditional expression that follows it. For example,</p>

<pre><code class="language-vb">If (myteam = "BYU") And Not (yourteam = "Utah" Or yourteam = "USU") Then
    MsgBox "We are going to get along OK."
End If
</code></pre>

<p>Notice that Not True --&gt; False and Not False --&gt; True.</p>
<h2>Evaluating Number Ranges</h2>
<p>When doing computations, it is common to need to determine if a number is inside a range. For example, in mathematics it is common two write a statement like this:</p>
<blockquote>
  <p>0 &le; x &le; 5</p>
</blockquote>
<p>When writing this as a compound conditional expression, it is tempting to write it as follows:</p>

<pre><code class="language-vb">If 0 <= x <= 5 Then
</code></pre>


<p>However, this is <strong>NOT</strong> logically equivalent to the statement shown above. For example, suppose that <strong>x = -10</strong>, which is outside the range and should make the expression evalue to False. The expression is evaluated in two parts from left to right, so the first part evaluated is <strong>0 &lt;= x</strong>, which returns a value of False. The value of False is then substituted for the first part of the expression and the remaining expression is then evaluated as <strong>False &lt;= 5</strong>. Whenever a boolean value (True/False) is compared to numerical value, <strong>True = 1</strong> and <strong>False = 0</strong>. Therefore, this expression is evaluated as <strong>0 &lt;= 5</strong>, which is True, leading to an incorrect result.</p>
<p>The proper way to write this expression in VB is:</p>

<pre><code class="language-vb">If 0 <= x And x <= 5 Then
</code></pre>


<p>In this case, the two sides are evaluated independently and then combined with the <strong>And</strong> operator, resulting in the correct answer.</p>
<h2>If Statements and Controls</h2>
<p>If statements are commonly used to determine the state of controls. Suppose you have a checkbox called <strong>chkResizeImage</strong>. You could check the state of the <strong>Value</strong> property as follows:</p>

<pre><code class="language-vb">If chkResizeImage.Value = True Then
</code></pre>

<p>Note that for a checkbox and option control, the Value property is a Boolean variable that equals True if the control is selected, and False otherwise. Since the Value property is the default property for each of these objects, you can simplify this statement by omitting the <strong>.Value</strong> part as follows:</p>


<pre><code class="language-vb">If chkResizeImage = True Then
</code></pre>


<p>Furthermore, this statement can be further simplified as follows:</p>

<pre><code class="language-vb">If chkResizeImage Then
</code></pre>


<p>In other words, the <strong>= True</strong> part is redundant because <strong>chkResizeImage = True</strong> is logically equivalent to the value of chkResizeImage (the expression is true when chkResizeImage is true).</p>

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
      <td> <strong>Score Keeper -</strong> Use an IF THEN expression to provide feedback on a calculated score.</td>
      <td align="center" valign="top">Easy</td>
      <td align="center" valign="top"><a href="score_keeper.xlsm">score_keeper.xlsm</a></td>
      <td align="center" valign="top"><a href="score_keeper_key.xlsm">score_keeper_key.xlsm</a></td>
    </tr>
    <tr>
      <td><strong>Density - </strong>Use an IF THEN expression with a check box and a conditional expression to display the correct density with the values given.</td>
      <td align="center" valign="top">Medium</td>
      <td align="center" valign="top"><a href="density.xlsm">density.xlsm</a></td>
      <td align="center" valign="top"><a href="density_key.xlsm">density_key.xlsm</a></td>
    </tr>
    <tr>
      <td><strong>Receipt -</strong> Create an IF THEN expression that takes a few conditions into consideration when generating a receipt for a customer.</td>
      <td align="center" valign="top">Hard</td>
      <td align="center" valign="top"><a href="receipt.xlsm">receipt.xlsm</a></td>
      <td align="center" valign="top"><a href="receipt_key.xlsm">receipt_key.xlsm</a></td>
    </tr>
  </tbody>
</table>
<br />

<?php 
require "../footer.php";
?>
</body>
</html>
