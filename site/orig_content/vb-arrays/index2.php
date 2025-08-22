<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Excel VBA Primer</title>
<link href="../../nljstyles.css" rel="stylesheet" type="text/css" />
<link href="../../prism/prism.css" rel="stylesheet" />
</head>

<body>
<script src="../../prism/prism.js"></script>

<h1>Using Arrays to Store Sets of Data</h1>
<p>In many applications (especially in engineering), it is necessary to work with variable containing a set of related data. For example, one may have a list (vector)  or matrix of numbers that need to be stored and used as part of some type of linear algebra calculations. Suppose you have a vector called &quot;x&quot; with ten double values. One solution would be to declare your variables like this:</p>
 <pre><code class="language-vb">Dim x1 As Double
Dim x2 As Double
Dim x3 As Double
Dim x4 As Double
Dim x5 As Double
Dim x6 As Double
Dim x7 As Double
Dim x8 As Double
Dim x9 As Double
Dim x10 As Double</code></pre> 

<p>Then in your code, would reference the variables independently:</p>
<pre><code class="language-vb">x1 = Cells(4,1)
x2 = Cells(4,2)
...
x8 = x7 * Sin(x6) + 2</code></pre>
<p>This works, but it is not an efficient approach, especially for cases where you have hundreds or thousands of items in the vector.</p>
<h2>One-Dimensional Arrays</h2>
<p>A more efficient way to handle the problem described above is to declare a one-dimensional array variable as follows:</p>
<blockquote>
  <p class="code"><span class="code_keyword">Dim</span> x(1 <span class="code_keyword">To</span> 10) <span class="code_keyword">As Double</span></p>
</blockquote>
<p>Then the elements of the array can be referenced like this:</p>
<blockquote>
  <p><span class="code">x(1) = Cells(4,1)<br />
    x(2) = Cells(4,2)<br />
    ...<br />
    x(8) = x(7) * Sin(x(6)) + 2</span></p>
</blockquote>
<p>One of the benefits of using arrays is that they can be easily traversed using the <strong>For i = ...</strong> style of loop. For example, the following code declares an array of 1000 doubles and initializes the value of each item in the array to = 1.5.</p>
<blockquote>
  <p class="code"><span class="code_keyword">Dim</span> v(1 <span class="code_keyword">To</span> 1000) <span class="code_keyword">As Double</span><br />
  <span class="code_keyword">Dim</span> i <span class="code_keyword">As Integer</span></p>
  <p class="code"><span class="code_keyword">For</span> i = 1 <span class="code_keyword">To</span> 1000<br />
  &nbsp;&nbsp;v(i) = 1.5<br />
  <span class="code_keyword">Next</span> i</p>
  <p class="code">&nbsp;</p>
 <pre><code class="language-vb">
'This is a comment
Dim v(1 To 1000) As Double
Dim i As Integer
  
For i = 1 To 1000
  v(i) = 1.5
Next i

</code></pre>
</blockquote>
<p>When declaring arrays, we can define the lower and uppper bounds of the array indices. For example,</p>
<blockquote>
  <p class="code"><span class="code_keyword">Dim</span> p(5 <span class="code_keyword">To</span> 15) <span class="code_keyword">As Integer</span></p>
  <p class="code">p(5) = 1<br />
    p(4) = 2<br />
    ...<br />
    p(14) = 12<br />
    P(15) = p(14) * p(12)
  </p>
</blockquote>
<p>It is also possible to omit the first part of the array bounds:</p>
<blockquote>
  <p class="code"><span class="code_keyword">Dim</span> x(100) <span class="code_keyword">As Double</span></p>
</blockquote>
<p>But if you do so, this declares an array of 101 items! The array bounds are from 0 to 100.</p>
<h2>Multi-Dimensional Arrays</h2>
<p>In some cases, it is useful to declare multi-dimensional arrays. For example, if we want to declare variable to store a matrix with 20 rows and 15 columns, we could do the following:</p>
<blockquote>
  <p class="code"><span class="code_keyword">Dim</span> n(1 <span class="code_keyword">To</span> 20, 1 <span class="code_keyword">To</span> 15) <span class="code_keyword">As Double</span></p>
</blockquote>
<p>And the individual elements of the arrays are referenced like this:</p>
<blockquote>
  <p class="code">m(1, 1) = 0.5<br />
    m(13, 5) = m(2, 12) * 2
  </p>
</blockquote>
<p>To traverse the elements of the array, we can use a nested set of <strong>For i = ...</strong> loops.</p>
<blockquote>
  <p class="code"><span class="code_keyword">Dim</span> i <span class="code_keyword">As Integer</span>, j <span class="code_keyword">As Integer</span></p>
  <p class="code"><span class="code_keyword">For</span> i = 1 <span class="code_keyword">To</span> 20<br />
    &nbsp;&nbsp;<span class="code_keyword">For</span> j  = 1 <span class="code_keyword">To</span> 15<br />
  &nbsp;&nbsp;&nbsp;&nbsp;m(i, j) = 1 / m(i, j)<br />
  &nbsp;&nbsp;<span class="code_keyword">Next</span> j<br />
  <span class="code_keyword">Next</span> i </p>
</blockquote>
<p>Arrays are not limited to two dimensions. You can use three, four, or more dimensions if you like.</p>
<blockquote>
  <p class="code"><span class="code_keyword">Dim</span> q(1 <span class="code_keyword">To</span> 2, 1 <span class="code_keyword">To</span> 100, 1 <span class="code_keyword">To</span> 50) <span class="code_keyword">As Double</span></p>
  <p class="code">q(1, 40, 25) = 4.5</p>
</blockquote>
<p>The maximum number of dimensions you can use is 32.</p>
<h2>Dynamic Arrays - The ReDim Statement</h2>
<p>When you declare an array using the Dim statement, the dimension bounds must be constants. However, there are cases when it is necessary to declare the bounds (especially the upper bound) using a variable. In other words, you many not know until you run your code how big the array needs to be. In this case, you should use the ReDim statement:</p>
<blockquote>
  <p class="code"><span class="code_keyword">Dim</span> n <span class="code_keyword">As Integer</span><br />
    ...<br />
    n = ....  <span class="code_comment">'&lt;--- Do something to determine how big the array needs to be</span><br />
    ...<br />
    <span class="code_keyword">ReDim</span> x(1 <span class="code_keyword">To</span> n) <span class="code_keyword">As Double</span><br />
    <span class="code_keyword">For</span> i = 1 <span class="code_keyword">To</span> n<br />
    &nbsp;&nbsp;x(i) = ...<br />
  <span class="code_keyword">Next</span> i </p>
</blockquote>
<p>A classic example involving spreadsheets is when we want to read a column of values into an array, but there may be any number of items in the column. We can either just guess what the maximum size could possibly be and declare a really big array, or we could write a short loop to count the number of items in the array and then declare an array just big enough to read the values. For example, consider the following range of cells in a worksheet:</p>
<blockquote>
  <p><img src="screenshot.jpg" width="261" height="632" alt=""/></p>
</blockquote>
<p>Suppose we wanted to read these two columns of numbers into two arrays, <strong>x</strong> and <strong>f</strong>, in order to peform some numerical calculations. We could declare the arrays to be 24 items long using the Dim statement and just allow for some empty space. Or we could loop through the cells first in order to find out exactly how many items are in the list and then declare two arrays to the exact size needed and then load the values from the cells into the array as shown here:</p>
<blockquote>
  <p><img src="x_fx_code.jpg" width="347" height="306" alt=""/></p>
</blockquote>
<p>At this point, the two arrays could be used directly in formulas and could be indexed from 1 to n.</p>
<p>Another case where the ReDim statement can be used is when you need to resize an array (make it larger or smaller) during run-time. For example,</p>
<blockquote>
  <p class="code">n = 1000<br />
    <span class="code_keyword">ReDim</span> x(1 <span class="code_keyword">To</span> n) <span class="code_keyword">As Double</span><br />
    ...<br />
    n = n + 100
    <br />
  <span class="code_keyword">ReDim</span> x(1 <span class="code_keyword">To</span> n) <span class="code_keyword">As Double</span></p>
</blockquote>
<p>After the last line, the array would contain 1100 items. If your array contains information that you do not want to lose, you should use the Preserve keyword.</p>
<blockquote>
  <p class="code">n = 1000<br />
    <span class="code_keyword">ReDim</span> x(1 <span class="code_keyword">To</span> n) <span class="code_keyword">As Double</span><br />
    ...<br />
    n = n + 100 <br />
    <span class="code_keyword">ReDim Preserve</span> x(1 <span class="code_keyword">To</span> n) <span class="code_keyword">As Double</span></p>
</blockquote>
<p>This ensures that the items stored in the array are not lost when the array is resized.<br />
</p>
<h2>Arrays vs. Ranges</h2>
<p>You many have noticed that a 2D array is very similar to a range: both have a rows and columns. In fact, when we use the Cells object to reference a set of cells using row and column indices like this:</p>
<blockquote>
  <p class="code">Cells(4, 5) = &quot;Hello&quot;</p>
</blockquote>
<p>or</p>
<blockquote>
  <p class="code"><span class="code_keyword">For</span> myrow = 6 <span class="code_keyword">To</span> 20<br />
    &nbsp;&nbsp;Cells(myrow, 2) = &quot;&quot;<br />
    <span class="code_keyword">Next</span> myrow</p>
</blockquote>
<p>We are essentially using the Cells object as a 2D array. </p>
<p>So the question arises, if you already have your data in a range of cells, why not just reference them directly using the Cells object as opposed to loading the values into a 1D or 2D array? In some cases, it is just a matter of personal preference. In other cases, the code can be much cleaner and easier to follow when you first load the values into arrays. For example, consider the Numeric Integration example in the previous section above. After loading the two columns into the x and f arrays, the calculations performed to numerically integrate the function defined by the numbers mimics the corresponding mathematical formula because the arrays are indexed from 1 to n. If one were to do the same calculations using the Cells object, it would be necessary to offset the row and column indices to match the location where the cells are stored on the worksheet. This makes the code much more difficult to write and debug.</p>
<p>&nbsp;</p>
<?php 
require "../footer.php";
?>
</body>
</html>
