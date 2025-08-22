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
<h1>Declaring and Using Variables</h1>
<p>Variables are a fundemental part of any programming language. A variable is a placeholder for a piece of information corresponding to a number, date, text string, etc. You can store information in a variable and then retrieve the information from the variable later as you use the variable as part of an expression or as part of an assignment statement that passes the information somewhere else. </p>
<h2>Simple Method</h2>
<p>The simplest way to use variables is to just reference them in your code. For example, consider the following code:</p>

<pre><code class="language-vb">Private Sub cmdCalculateArea_Click()
Radius = 4.3
Pi = 3.14159
Area = Pi * Radius ^ 2
Circ = 2 * Pi * Radius
MsgBox "Area = " & Area
MsgBox "Circumference = " & Circ
End Sub
</code></pre>


<p>This code computes the area and cicumference of a circle with radius of 4.3. We first assign values to the <strong>Radius</strong> and <strong>Pi</strong> variables and then use those valuables to calculate the area and circumference and the store these two values in the <strong>Area</strong> and <strong>Circ</strong> variables. The values stored in the Area and Circ variables are then displayed using the MsgBox calls. This code brings up the following two messages:</p>
<blockquote>
  <p><img src="msgbox1.png" width="161" height="154" alt=""/> <img src="msgbox2.png" width="209" height="154" alt=""/></p>
</blockquote>
<p> As you can see, variables provide a simple and convenient way to perform calculations.</p>
<h2>Declaring Variables</h2>
<p>While the approach shown above is  simple, it can be dangerous. For example, consider the following code:</p>


<pre><code class="language-vb">Private Sub cmdCalculateArea_Click()
Radias = 4.3
Pi = 3.14159
Area = Pi * Radius ^ 2
Circ = 2 * Pi * Radius
MsgBox "Area = " & Area
MsgBox "Circumference = " & Circ
End Sub
</code></pre>

<p>Note the difference on the very first list of code (&quot;Radias = ...&quot;). The word &quot;Radius&quot; has been mispelled, but it is spelled correctly in the rest of the code. This is what the code produces:</p>
<blockquote>
  <p><img src="msgbox1b.png" width="154" height="154" alt=""/> <img src="msgbox2b.png" width="165" height="154" alt=""/></p>
</blockquote>
<p>which is clearly wrong! What happend? Whenever the VB compiler encounters a variable name it does not recognize, it  creates a new variable with that name and initializes the value to zero. So there is a variable called &quot;Radias&quot; with a value of 4.3, but when the lines for calculating the area and circumference were encountered, the compiler created a new variable called &quot;Radius&quot; with a default value of zero and used that variable in the calculations. That is why the area and circumference ended up with a value of zero. This is a logic error and can be extremely difficult to track down and fix in some cases.</p>
<p>It is easy to make spelling mistakes, so this can lead to lots of bugs in your code. Fortunately, there is a simple way to ensure that this never happens. First of all, you should put the following line at the top of each of your code modules:</p>

<pre><code class="language-vb">Option Explicit
</code></pre>

<p>This line forces you to declare all variables before you use them. You declare variables using a <strong>Dim</strong> statement as follows:</p>

<pre><code class="language-vb">Dim radius, pi, area, circ
</code></pre>


<p>This statement defines the variables you intend to use and the compiler immediately creates the variables in the list when this statement is encountered. Then if you ever encounter a mispelled variable, the code stops with a run-time error and you get a message like this:</p>
<blockquote>
  <p><img src="compile_error.png" width="553" height="251" alt=""/></p>
</blockquote>
<p>It is then a simple matter to correct the spelling and fix the bug. Therefore, you should ALWAYS remember to include the Option Explicit statement when using variables. It will prevent a huge amount of nasty bugs.</p>
<h2>Variable Types</h2>
<p>When we create variables to use in our code, we typically have a specific objective in mind and we plan to store a specific type of data in each variable. We can change our Dim statement to more explicitly define what type of variable we are declaring as follows:</p>

<pre><code class="language-vb">Dim n As Integer, name As String, birthday As Date
</code></pre>

<p>The advantage of explicitly declaring types is that if you ever try to store something of one type in a variable of another type, you will get an error message. Here are the most commonly-used variable types in VBA:</p>
<blockquote>
  <table width="876" border="1">
    <tr>
      <td width="342"><strong>Type</strong></td>
      <td width="57" align="center"><strong>Mem</strong></td>
      <td width="463" align="center"><strong>Range</strong></td>
    </tr>
    <tr>
      <td>Boolean</td>
      <td align="center">1B</td>
      <td align="center">True/False</td>
    </tr>
    <tr>
      <td>Currency</td>
      <td align="center">8B</td>
      <td align="center">-922,337,203,685,477.5808 to 922,337,203,685,477.5807</td>
    </tr>
    <tr>
      <td>Date</td>
      <td align="center">8B</td>
      <td align="center">January 1, 100 to December 31, 9999</td>
    </tr>
    <tr>
      <td>Single (single precision floating point numbers)</td>
      <td align="center">4B</td>
      <td align="center">+/-1e-45 to +/-1e+38</td>
    </tr>
    <tr>
      <td>Double (double precision floating point numbers)</td>
      <td align="center">8B</td>
      <td align="center">+/-1e-324 to +/-1e+308</td>
    </tr>
    <tr>
      <td>Integer (single precision whole numbers)</td>
      <td align="center">2B</td>
      <td align="center">-32,768 to 32,767</td>
    </tr>
    <tr>
      <td>Long (double precision whole numbers)</td>
      <td align="center">4B</td>
      <td align="center">-2,147,483,648 to 2,147,483,647</td>
    </tr>
    <tr>
      <td>String</td>
      <td align="center">Varies</td>
      <td align="center">0 to ~2 billion  characters</td>
    </tr>
    <tr>
      <td>Variant</td>
      <td align="center">Varies</td>
      <td align="center">All of the above</td>
    </tr>
  </table>
  <br />
<em>Source: <a href="http://msdn.microsoft.com/en-us/library/aa263420(v=vs.60).aspx">MSDN - Visual Basic for Applications Reference, Data Type Summary</a> </em></blockquote>
<p>Singles and doubles both work well for floating point numbers (numbers with both a whole and decimal part, Ex. -207.393). However, a single only preserves about 7 digits of accuracy while a double preserves about 15 digits. Therefore we typically use doubles just to be safe.</p>
<p>The Variant type is a special case and is somewhat unique to VB. A variant can hold any type of information. When you store something in a Variant, part of the Variant memory is used to mark the type of data current stored in the variable. When you access the data in the variable, the data type information is used to process it properly. Therefore, a Variant is kind of a general purpose type. You can explicitly declare something to be a Variant as follows:</p>

<pre><code class="language-vb">Dim x As Variant, y As Variant, z As Variant
</code></pre>


<p>Or you can simply do this:</p>

<pre><code class="language-vb">Dim x, y, z
</code></pre>


<p>and you get the same result. In other words, if you don't explicitly declare the type for a variable, it defaults to Variant. You need to be careful with this method. For example, many people assume that this statement:</p>

<pre><code class="language-vb">Dim x, y, z As Double
</code></pre>

<p>declares all three variables as Doubles. However, that is not true. The only variable that ends up as a Double is <strong>z</strong>. Since no type was defined for <strong>x</strong> and <strong>y</strong>, they default to Variants. In other words, it is the same as doing the following:</p>

<pre><code class="language-vb">Dim x As Variant, y As Variant, z As Double
</code></pre>

<h2>Default Values</h2>
<p>When you declare a numeric variable, the default value of the variable is always zero. For example, consider the following code:</p>

<pre><code class='language-vb'>Dim n As Integer

Range("B4") = n
n = n + 1
Range("B5") = n</code></pre>

<p>After this code is executed, there is a value of <b>0</b> in cell <b>B4</b> and a value of <b>1</b> in cell <b>B5</b>. If you reference a variable after it is declared but before assigning a value to it, it contains the default value of zero.</p>
<p>Likewise, the default value for a string variable = "" (empty string). And the default value for a boolean variable = False.</p>
<p>If you are ever in doubt, there is no harm in assigning a value to a variable before using it.</p>
<h2>Variable Names</h2>
<p>When coming up with names for your variables, a few simple rules must be follows. </p>
<ol>
  <li>You can only use the characters a-z, A-Z, 0-9, and the underscore character (&quot;_&quot;).</li>
  <li>You cannot use reserved VB words such as &quot;If&quot;, &quot;Dim&quot;, etc.</li>
</ol>
<p>You should also try to make your variable names reflective of the variable usage when follows. For example, use &quot;last_name&quot;, &quot;first_name&quot; rather than &quot;x1&quot;, &quot;x2&quot; if you are going to store names in your variables.</p>
<h2>Assignment Statements</h2>
<p>When working with variables, it is important to understand how assignment statements work. To store a value in a variable, you would do something like this:</p>

<pre><code class="language-vb">x = 283.922
</code></pre>

<p>As a mathematical expression, this statement would be interpreted as &quot;x is equal to 283.922&quot;. However, that is not the best way to interpret this statement in programming logic. When used as a stand-alone statement, it literally means &quot;take the value of 283.922 and store it in the variable called x&quot;. This is why we call this an assignment statement. Here is another example:</p>

<pre><code class="language-vb">x = 2
y = 8
x = y 
</code></pre>


<p>As a sequence of mathematical expressions this does not make sense because 2 does not equal 8. But as an assignment statement, the last line means &quot;Take the current value of y (which is 8) and store it in x.&quot; After the assignment statement is completed, both x and y have a value of 8. In some cases, the thing on the right side is an expression:</p>

<pre><code class="language-vb">x = (7 * y) / 3
</code></pre>


<p>In this case we evaluate the expression on the right using the current value of y and then store the result in x. It should be noted that when writing an assignment statement, the item on the left side must be a variable or object that can store the result of the expression on the right. In other words, this would not make sense:</p>

<pre><code class="language-vb">(y * y) / 3 = x '<-- WRONG!
</code></pre>

<p>Here is another example. In this case we are going read the value from cell A4 and store it in x.</p>

<pre><code class="language-vb">x = Range("A4")
</code></pre>

<p>Once again, the order is critical. If what you really wanted to do was take the value of x and store it in cell A4, you would reverse the statement as follows:</p>

<pre><code class="language-vb">Range("A4") = x
</code></pre>

<p>This illustrates that cells are similar to variables. You can store values in cells and then retrieve those values later.</p>
<p>In summary, the best way to think of an assignment statement is <strong>&quot;Take the value of the expression on the right, and store it in the variable/object on the left.&quot;</strong></p>
<h2>Constants</h2>
<p>In some cases it is useful to declare a variable that never changes. Such a variable is called a constant. Here is a commonly used constant:</p>

<pre><code class="language-vb">Const pi As Double = 3.14159
</code></pre>


<p>After making this declaration, you can use pi anywhere in your code as an alias for 3.15159. For example,</p>

<pre><code class="language-vb">area = pi * radius ^ 2
cicumf = 2 * pi * r 
</code></pre>

<p>You could do the same thing by declaring pi as a variable (see code at the top of this page). The advantage of using a constant is that it ensures that you don't accidentally change the value of the constant. If you attempt to do so, you will get an error message.</p>
<h2>Scope</h2>
<p>When dealing with variables, we often need to consider something called &quot;scope&quot;. A variable's scope defines in which part of the code the variable can be utilized. In most cases, we define a variable inside of a sub or function as follows:</p>

<pre><code class="language-vb">Sub mysub()
Dim x As Double, y As Double, sum As Double
x = Range("A4")
y = Range("B4")
sum = x + y
Range("C4") = sum
End Sub
</code></pre>


<p>In this case, the scope of the variables x, y, and sum is within mysub. If you were to use a variable with the same name(s) in another sub as follows:</p>

<pre><code class="language-vb">Sub mysub()
Dim x As Double, y As Double, sum As Double
x = Range("A4")
y = Range("B4")
sum = x + y
Range("C4") = sum
End Sub

Sub anothersub()
Dim x As Double, y As Double, sum As Double
x = 5
y = 12
Range("C4") = y * x
End Sub
</code></pre>

<p>each of the two subs would be referencing two different sets of variables. However, if you declare the variables outside the subs like this:</p>

<pre><code class="language-vb">Dim x As Double, y As Double, sum As Double

Sub mysub()
x = Range("A4")
y = Range("B4")
sum = x + y
Range("C4") = sum
End Sub

Sub anothersub()
x = 5
y = 12
Range("C4") = y * x
End Sub
</code></pre>

<p>then the scope of the variables is defined as the entire module and both subs would be referencing the same variables. If you want a variable to be shared by every sheet, form, or module in the entire project, you can declare it as a global variable as follows:</p>

<pre><code class="language-vb">Global x As Double, y As Double, sum As Double
</code></pre>

<p>Globals can only be defined inside a module (you can't do this in a sheet or form). </p>
<p>While it may be tempting to declare variables outside of subs or as globals, this is generally viewed as poor programming practice. In the vast majority of cases, it is cleaner and less confusing if you declare your variables inside each sub or function where they are used. If you need to pass information from one sub to another, do it via  input arguments.</p>

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
      <td> <strong>Total Head -</strong> Pass cell values into variables and use the variables to calculate the total fluid head.</td>
      <td align="center" valign="top">Easy</td>
      <td align="center" valign="top"><a href="total_head.xlsm">total_head.xlsm</a></td>
      <td align="center" valign="top"><a href="total_head_key.xlsm">total_head_key.xlsm</a></td>
    </tr>
    <tr>
      <td><strong>Lengths, Dot Product - </strong>Use variables to store vectors and calculate their lengths and dot product.</td>
      <td align="center" valign="top">Medium</td>
      <td align="center" valign="top"><a href="lengths_dot_product.xlsm">lengths_dot_product.xlsm</a></td>
      <td align="center" valign="top"><a href="lengths_dot_product_key.xlsm">lengths_dot_product_key.xlsm</a></td>
    </tr>
    <tr>
      <td><strong>Citation Generator -</strong> Store the values/words of a resource as variables to use in the string generator that displays the correct citation.</td>
      <td align="center" valign="top">Hard</td>
      <td align="center" valign="top"><a href="citation_generator.xlsm">citation_generator.xlsm</a></td>
      <td align="center" valign="top"><a href="citation_generator_key.xlsm">citation_generator_key.xlsm</a></td>
    </tr>
  </tbody>
</table>
<br />

<?php 
require "../footer.php";
?>
</body>
</html>
