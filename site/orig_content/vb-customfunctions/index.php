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
<h1>Custom Functions</h1>
<p>One of the easiest ways to take advantage of VBA in Excel is to write custom 
  functions. Excel has a large number of built-in functions that you can use 
  in spreadsheet formulas. Examples include:</p>
<blockquote>
  <table width="556" border="0">
    <tr>
      <td width="144">=Average(&quot;A4:A20&quot;)</td>
      <td width="402"><i>Returns the average value in a range of cells</i></td>
    </tr>
    <tr>
      <td>=Sum(&quot;A4:A20&quot;)</td>
      <td><i>Returns the sum of a range of cells</i></td>
    </tr>
    <tr>
      <td>=Cos(0.34)</td>
      <td><i>Returns the cosine of a number</i></td>
    </tr>
  </table>
</blockquote>
<p>In general, a function takes one or more objects (values, ranges, etc.) as 
  input and returns a single object (typically a value). The things that are 
  sent to a function as input are called <b>arguments</b> and the thing that is 
  returned by a function is often called the <b>return value</b>.</p>
<p>In some cases, we may encounter situations where we need a function to do 
  something but the function is not provided by Excel. We can easily fix 
  this problem by creating a custom function in VBA.</p>
<h2>Syntax</h2>
<p>The basic syntax for a custom function is as follows:</p>
<pre><code class="language-vb">[Public] Function function_name(args) As Type
...
function_name = ...
...
End Function
</code></pre>


<p>The <b>Public</b> statement is optional. This means that the function 
  can be called by VB code outside the module where the function is declared and 
  from Excel formulas in cells. If you omit the <b>Public</b> statement, the 
  function is public by default (as opposed to <b>Private</b>).</p>
<p>The <b>function_name</b> is the name you provide to the function. This 
  can be any name you want, as long as you follow the same rules we use for 
  defining VB variable names.</p>
<p>The <b>args</b> are the arguments to the function. You can have any 
  number of arguments. The arguments are listed in the same way you declare 
  variables, except that you omit the <b>Dim</b> part. The <b>args</b> list 
  serves two purposes: 1) it defines how many arguments are used, and 2) it 
  defines the type for each argument. The following are some sample argument 
  lists:</p>
  
<pre><code class="language-vb">(x As Double, n As Double)

(r As Range)

(str1 As String, str2 As String, num As Integer)
</code></pre>  
  

<p>The <b>Type</b> part defines the type of object returned by the function. Typical examples are Double, Integer, String, and Boolean.</p>
<p>Somewhere in the code, you must have line where you set the <b>function name</b> equal to a value. You should think of the function name as a variable. You must store the value returned by the function in the variable at some point 
  before you hit the <b>End Function</b> statement.</p>
<p>There is one more important point, whenever you create a function that you 
  want to use in an Excel formula, it should always be placed in a module under 
  the <b>Modules</b> folder in the VBE.</p>
<h2><a name="examples" id="examples"></a>Examples</h2>
<p>Now let's look at some examples. The following function takes two 
  numbers as arguments and returns the minimum of the two numbers. This 
  basically duplicates the Min function provided by Excel, but it serves as a 
  useful example:</p>

<pre><code class="language-vb">Function my_min(a As Double, b As Double) As Double

If a < b Then
   my_min = a
Else
   my_min = b
End If

End Function
</code></pre>


<p>Once this function is created, you can then use it in one of your Excel 
  formulas as follows:</p>
  
<pre><code class="language-vb">=my_min(A5, B7)
=my_min(Sum(C3:C10), 0)
</code></pre>

<p>If you want to be lazy and not worry about declaring your types, you can simplify the first line of your function declaration as follows:</p>


<pre><code class="language-vb">Function my_min(a, b)

If a < b Then
   my_min = a
Else
   my_min = b
End If

End Function
</code></pre>


<p>Compare to the same function shown above. In this case, the input arguments and the return type are set to <strong>Variant</strong> by default. Both of these methods will work, but the advantage of declaring the types is that if you pass something to the function that is of the wrong type, you will get an error message.</p>
<p>Now let's look at something a little more complicated. In many cases, 
  we want our function to use a cell range as one of the arguments. The 
  following function returns the number of negative values in a range:</p>

<pre><code class="language-vb">Function num_neg(r As Range) As Integer

Dim c As Variant

For Each c In r
    If c.Value < 0 Then
		num_neg = num_neg + 1
    End If
Next c

End Function
</code></pre>


<p>The function could then be called from an Excel formula as follows:</p>

<pre><code class="language-vb">=num_neg(A5:B7)
</code></pre>


<p>The next function takes two arguments: a range and an integer n. It 
  computes the sum of values in the range minus the lowest n values. This 
  function takes advantage of the standard Excel functions <b>Sum</b> and <b>Small</b>. The Small function returns the n<sup>th</sup> lowest value in a range.</p>
  
<pre><code class="language-vb">Function dropsum(r As Range, n As Integer) As Double

Dim i As Integer

dropsum = Application.WorksheetFunction.Sum(r)

For i = 1 To n
   dropsum = dropsum - Application.WorksheetFunction.Small(r, i)
Next i

End Function
</code></pre>  
  
<p>This function could then be used in an Excel formula as follows:</p>

<pre><code class="language-vb">=dropsum(A5:B7, C10)
=dropsum(A5:B7, 5)</code></pre>


<h2>Functions in VB</h2>
<p>Finally, it should be noted that you can call custom functions from other 
  places in your VB code as well as from Excel formulas. For example, you 
  could use the my_min function defined above as follows:</p>
  
<pre><code class="language-vb">Dim x As Double
Dim y As Double
Dim z As Double

x = ...
y = ...
...
z = my_min(x, y)
...</code></pre>
  

<h2>Arguments - ByVal vs. ByRef</h2>
<p>If you call a VB function from someone else in your VB code, you need to be careful how you handle the arguments. In most cases, the arguments are used as input values to our computations and we don't attempt to change the values of the arguments. But if you do change the values of the arguments, you need to be understand what happens. For example, consider the following code.</p>

<pre><code class="language-vb">Function foo(x As Double) As Double

x = x - 1
foo = 2*x

End Function

Sub mysub()

Dim p As Double
Dim r As Double

p = 5
r = foo(p)

MsgBox p

End Sub</code></pre>

<p>Note that <strong>p</strong> was passed as an argument to the foo function where it was referenced as the argument <strong>x</strong>. In the function, <strong>x</strong> is decremented by 1. The value of <strong>p</strong> is then displayed using <strong>MsgBox</strong> in the sub after the function call is complete. The following is displayed:</p>
<p><img src="byval1.png" width="154" height="154" alt=""/></p>
<p>In other words, when we change the value of <strong>x</strong> in the function, we are simultaneously changing the value of <strong>p</strong>! Or put another way, <strong>x</strong> becomes an alias for <strong>p</strong> and any change we make to <strong>x</strong> is made to <strong>p</strong>. They are two names for the same thing. This can lead to unexpected consequences. We can isolate the <strong>x</strong> from <strong>p</strong> by adding &quot;<strong>ByVal</strong>&quot; in front of the argument as follows:</p>
<pre><code class="language-vb">Function foo(ByVal x As Double) As Double

x = x - 1
foo = 2*x

End Function</code></pre>
<p>In this case, after running <strong>mysub</strong>, the following is displayed:</p>
<p><img src="byval2.png" width="154" height="154" alt=""/></p>
<p>The <strong>ByVal</strong> qualifier forces the argument to be passed &quot;by value&quot;, meaning that <strong>x</strong> then becomes a copy of <strong>p</strong> and any changes we make to <strong>x</strong> are not reflected in p<strong>.</strong> The alternate qualified is <strong>ByRef</strong>, which indicates that the argument is directly linked to the variable passed in when the function is called. If you omit either qualifier, VBA defaults to <strong>ByRef</strong>, which is why our original example behaved the way it did.</p>
<h2>Rules for Functions</h2>
<p>There are a few simple rules that should be followed when writing custom functions.</p>
<p>1) When the function is called, arguments can be constants, values from cells, values from mathematical expressions, etc. It is not always from a cell, and it is most certainly not always from the same cell. For example, here are several fully legal ways to call the my_min function described above:
</p>

<pre><code class="language-vb">=my_min(A5,	B4)
=my_min(-4, 20.4)
=my_min(0, -B4/(A2 + 5))
</code></pre>


<p>In other words, your VBA code should never assume that the values are from cells. The values could come from any kind of expression.</p>
<p>2) Do not read directly from a cell inside a function. All input should be from the list of arguments. For example, consider the following function code:</p>

<pre><code class="language-vb">Function my_min2(a As Double, b As Double) As Double

a = Range("B4")
b = Range("B5")

If a < b Then
   my_min2 = a
Else
   my_min2 = b
End If

End Function
</code></pre>

<p>Note that the two arguments <strong>a</strong> and <strong>b</strong> are immediately reset using the values from cells B4 and B5. This is a very common programming error but it is <strong>WRONG</strong>! Functions are supposed to be general purpose in nature. You should be able to use your function in a formula in cells located anywhere on your worksheet. The problem with the code above is that no matter what you pass in as arguments, it will always use the values from <strong>B4</strong> and <strong>B5</strong>. You should never change the values of the input arguments in your code. If you want to apply the function to B4 and B5, put this in a cell formula after you write the function code:</p>

<pre><code class="language-vb">=my_min2(B4, B5)
</code></pre>


<p>3) Do not write directly to a cell from a function. In some cases, this will generate a calculate event, putting your spreadsheet into an infinite loop. The function should only do one thing: return a value. For example, consider the following code:
</p>

<pre><code class="language-vb">Function my_min3(a As Double, b As Double) As Double

If a < b Then
   my_min3 = a
Else
   my_min3 = b
End If

Range("C8") = my_min3

End Function
</code></pre>


<p>Once again, this is a common error, but it is <strong>WRONG</strong>! If you want to use the function to put a value in cell C8, use the function in a formula in cell C8 and the let the return value from the function generate the answer. The only result of a function should be the return value. If you want to write some code that changes values of one or more cells directly from the code, use a custom sub, not a function.</p>

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
      <td> <strong>Load Function -</strong> Create a custom function to calculate the excess load from a table of test data.</td>
      <td align="center" valign="top">Easy</td>
      <td align="center" valign="top"><a href="load_function.xlsm">load_function.xlsm</a></td>
      <td align="center" valign="top"><a href="load_function_key.xlsm">load_function_key.xlsm</a></td>
    </tr>
    <tr>
      <td><strong>Total Stress - </strong>Use a custom function to determine the total stress vs depth of a soil profile.</td>
      <td align="center" valign="top">Medium</td>
      <td align="center" valign="top"><a href="total_stress.xlsm">total_stress.xlsm</a></td>
      <td align="center" valign="top"><a href="total_stress_key.xlsm">total_stress_key.xlsm</a></td>
    </tr>
    <tr>
      <td><strong>Total Headloss -</strong> Create custom functions to calculate parts of the total headloss equation and then solve for total headloss.</td>
      <td align="center" valign="top">Hard</td>
      <td align="center" valign="top"><a href="total_headloss.xlsm">total_headloss.xlsm</a></td>
      <td align="center" valign="top"><a href="total_headloss_key.xlsm">total_headloss_key.xlsm</a></td>
    </tr>
  </tbody>
</table>
<br />

<?php 
require "../footer.php";
?>
</body>
</html>
