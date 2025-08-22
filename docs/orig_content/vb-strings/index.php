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
<h1>Using String Functions</h1>
<p>A string is a sequence of one or more text characters. A string constant is a sequence of characters in code delimited by double quotes (&quot;Hello world&quot;). A string variable is a variable used to store text. String variables are dynamically allocated to accomodate the amount of text you store in each string. They can hold up to two billion characters!</p>
<h2>Concatenation</h2>
<p>Recall that we can combine (concatenate) two or more strings using the <strong>&amp;</strong> operator. For example, the code:</p>
<pre><code class="language-vb">Dim str1 As String, str2 As String
str1 = "Enter to learn"
str2 = "go forth to serve"
MsgBox str1 & ", " & str2
</code></pre>
<p>Would produce:</p>
<blockquote>
  <p><img src="msgbox-motto.png" width="230" height="154" alt=""/></p>
</blockquote>
<p>Note that this example involves both string variables (str1, str2) and string constants (&quot;, &quot;).</p>
<h2>String Functions</h2>
<p>In many cases, we need to manipulate string constants and variables. VB provides a rich suite of functions for manipulating strings. Some of the more widely used string functions are as follows:</p>
<table border="1" cellpadding="3" cellspacing="2" style="border-collapse: collapse" width="86%" id="table1">
  <tr>
    <td width="170"><b>Function</b></td>
    <td width="360"><b>Description</b></td>
    <td width="336"><strong>Example</strong></td>
    <td width="123" align="center"><strong>Result</strong></td>
  </tr>
  <tr>
    <td width="170" valign="top">Left(str, numchar)</td>
    <td width="360" valign="top">Extracts the specified number of characters from the 
      left side of the string <b>str</b>.</td>
    <td width="336" valign="top">mystr = Left(&quot;Hello World&quot;, 5)</td>
    <td width="123" align="center" valign="top">&quot;Hello&quot;</td>
  </tr>
  <tr>
    <td width="170" valign="top">Right(str, numchar)</td>
    <td width="360" valign="top">Extracts the specified number of characters from the 
      right side of the string <b>str.</b></td>
    <td width="336" valign="top">mystr = Right(&quot;Hello World&quot;, 5)</td>
    <td width="123" align="center" valign="top">&quot;World&quot;</td>
  </tr>
  <tr>
    <td width="170" valign="top">Mid(str, startchar, [numchar])</td>
    <td width="360" valign="top">Extracts a string of length <b>numchar</b> from the 
      middle of <b>str</b>, starting at <b>startchar</b>. If numchar is 
      omitted, the entire right-hand portion of the string, beginning at 
      startchar, is extracted.</td>
    <td width="336" valign="top"><p>mystr = Mid(&quot;Hello World&quot;, 7, 1)<br />
      mystr = Mid(&quot;Hello World&quot;, 7)</p></td>
    <td width="123" align="center" valign="top"><p>&quot;W&quot;<br />
      &quot;World&quot;</p></td>
  </tr>
  <tr>
    <td width="170" valign="top">Len(str)</td>
    <td width="360" valign="top">Returns the length of <b>str</b>.</td>
    <td width="336" valign="top">n = Len(&quot;Hello World&quot;)</td>
    <td width="123" align="center" valign="top">11</td>
  </tr>
  <tr>
    <td width="170" valign="top">UCase(str)</td>
    <td width="360" valign="top">Converts the string to all uppercase characters.</td>
    <td width="336" valign="top">mystr = UCase(&quot;Hello World&quot;)</td>
    <td width="123" align="center" valign="top">&quot;HELLO WORLD&quot;</td>
  </tr>
  <tr>
    <td width="170" valign="top">LCase(str)</td>
    <td width="360" valign="top">Converts the string to all lowercase characters.</td>
    <td width="336" valign="top">mystr = LCase(&quot;Hello World&quot;)</td>
    <td width="123" align="center" valign="top">&quot;hello world&quot;</td>
  </tr>
  <tr>
    <td width="170" valign="top">LTrim(str)</td>
    <td width="360" valign="top">Returns a copy of the string without leading spaces</td>
    <td width="336" valign="top">mystr = LTrim(&quot;  Hello  &quot;)</td>
    <td width="123" align="center" valign="top">&quot;Hello  &quot;</td>
  </tr>
  <tr>
    <td width="170" valign="top">RTrim(str)</td>
    <td width="360" valign="top">Returns a copy of the string without trailing spaces</td>
    <td width="336" valign="top">mystr = RTrim(&quot;  Hello  &quot;)</td>
    <td width="123" align="center" valign="top">&quot;  Hello&quot;</td>
  </tr>
  <tr>
    <td width="170" valign="top">Trim(str)</td>
    <td width="360" valign="top">Returns a copy of the string without leading or 
      trailing spaces</td>
    <td width="336" valign="top">mystr = Trim(&quot;  Hello  &quot;)</td>
    <td width="123" align="center" valign="top">&quot;Hello&quot;</td>
  </tr>
  <tr>
    <td width="170" valign="top">StrReverse(str)</td>
    <td width="360" valign="top">Returns a copy of the string in reverse order</td>
    <td width="336" valign="top">mystr = StrReverse(&quot;Hello World&quot;)</td>
    <td width="123" align="center" valign="top">&quot;dlroW olleH&quot;</td>
  </tr>
  <tr>
    <td width="170" valign="top">Replace(expression, find, replace, [start], [count], 
      [compare])</td>
    <td width="360" valign="top">Replaces each instance of “<b>find</b>” with “<b>replace</b>” 
      in “<b>expression</b>”.</td>
    <td width="336" valign="top">mystr = Replace(&quot;Hello World, &quot;Hello&quot;, &quot;Goodbye&quot;)<br />
      mystr = Replace(&quot;Hello World, &quot;l&quot;, &quot;&quot;)<br />
      mystr = Replace(&quot;Hello World, &quot;o&quot;, &quot;&quot;, 1, 1)<br />
      mystr = Replace(&quot;Hello World, &quot;o&quot;, &quot;&quot;, 1) <br />
      mystr = Replace(&quot;Hello World, &quot;l&quot;, &quot;&quot;, 6) </td>
    <td width="123" align="center" valign="top">&quot;Goodbye World&quot;<br />
      &quot;Heo Word&quot;<br />
      &quot;Hell World&quot; <br />
      &quot;Hell Wrld&quot;<br />
      &quot;Hello Word&quot; </td>
  </tr>
  <tr>
    <td width="170" valign="top">InStr([start], string1, string2, [compare])</td>
    <td width="360" valign="top">Returns an integer representing the 
      position of <b>string2</b> inside <b>string1</b>.&nbsp; <b>Start</b> is 
      an optional starting location (if omitted, search starts at position 1).&nbsp; 
      If string2 is not found in string1, the function returns zero.</td>
    <td width="336" valign="top">n = InStr(1, &quot;Hello World&quot;, &quot;W&quot;)<br />
      n = InStr(1, &quot;Hello World&quot;, &quot;N&quot;)</td>
    <td width="123" align="center" valign="top">7<br />
      0</td>
  </tr>
  <tr>
    <td valign="top">StrConv(string, conversion, [LCID])</td>
    <td valign="top">Returns a copy of <strong>string</strong> after modifying the string based on the conversion argument. The <strong>conversion</strong> argument should be a vb constant  with options including vbUpperCase, vbLowerCase, vbProperCase.</td>
    <td valign="top">mystr = StrConv(&quot;Hello World&quot;, vbUpperCase)<br />
    mystr = StrConv(&quot;Hello World&quot;, vbLowerCase)<br />
    mystr = StrConv(&quot;hello world&quot;, vbProperCase)</td>
    <td align="center" valign="top">HELLO WORLD<br />
      hello world<br />
      Hello World</td>
  </tr>
</table>
<p>These function can be combined in creative ways to achieve a variety of results. For example, consider the following table:</p>
<blockquote>
  <p><img src="name_table1.png" width="276" height="129" alt=""/></p>
</blockquote>
<p>Suppose we wanted to populate the third column with full names. We could do that with the following code:</p>

<pre><code class="language-vb">Dim first As String
Dim last As String
Dim full As String
For myrow = 2 To 5
    first = Cells(myrow, 2)
    last = Cells(myrow, 1)
    full = first & " " & last
    Cells(myrow, 3) = full
Next myrow
</code></pre>

<p>Resulting in:</p>
<blockquote>
  <p><img src="name_table2.png" width="287" height="157" alt=""/></p>
</blockquote>
<p>If we wanted the full name to be &quot;last, first&quot; format, we could restructure our code as follows:</p>

<pre><code class="language-vb">For myrow = 2 To 5
    first = Cells(myrow, 2)
    last = Cells(myrow, 1)
    full = last & ", " & first
    Cells(myrow, 3) = full
Next myrow
</code></pre>

<p>Resulting in:</p>
<blockquote>
  <p><img src="name_table3.png" width="291" height="135" alt=""/></p>
</blockquote>
<p>Now suppose we wanted the full name to be in all caps and we want to ensure that all leading and trailing spaces (if any) are removed. We could add references to string functions as follows:</p>

<pre><code class="language-vb">For myrow = 2 To 5
    first = Cells(myrow, 2)
    last = Cells(myrow, 1)
    full = UCase(Trim(last) & ", " & Trim(first))
    Cells(myrow, 3) = full
Next myrow
</code></pre>

<p>Resulting in:</p>
<blockquote>
  <p><img src="name_table4.png" width="301" height="148" alt=""/></p>
</blockquote>
<h2>Sample Code</h2>
<p>The workbook associated with the examples on this page can be downloaded here:</p>
<p><a href="sample code.xlsm">sample code.xlsm</a></p>

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
      <td> <strong>Citation -</strong> From a range of values, create a citation as a string and display the result as a message box.</td>
      <td align="center" valign="top">Easy</td>
      <td align="center" valign="top"><a href="citation.xlsm">citation.xlsm</a></td>
      <td align="center" valign="top"><a href="citation_key.xlsm">citation_key.xlsm</a></td>
    </tr>
    <tr>
      <td> <strong>Concatenation -</strong> By using concatenation, create a madlib generator.</td>
      <td align="center" valign="top">Easy</td>
      <td align="center" valign="top"><a href="concatenation.xlsm">concatenation.xlsm</a></td>
      <td align="center" valign="top"><a href="concatenation_key.xlsm">concatenation_key.xlsm</a></td>
    </tr>
    <tr>
      <td><strong>Secret Code - </strong>Use string functions to create a secret code that only you will know how to decipher.</td>
      <td align="center" valign="top">Medium</td>
      <td align="center" valign="top"><a href="secret_code.xlsm">secret_code.xlsm</a></td>
      <td align="center" valign="top"><a href="secret_code_key.xlsm">secret_code_key.xlsm</a></td>
    </tr>
    <tr>
      <td><strong>URL Deconstruct -</strong> Use a variety of string functions to deconsturct website URL addresses into their respective parts.</td>
      <td align="center" valign="top">Hard</td>
      <td align="center" valign="top"><a href="url_deconstruct.xlsm">url_deconstruct.xlsm</a></td>
      <td align="center" valign="top"><a href="url_deconstruct_key.xlsm">url_deconstruct_key.xlsm</a></td>
    </tr>
  </tbody>
</table>
<br />

<?php 
require "../footer.php";
?>
</body>
</html>
