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
<h1>Working with Objects</h1>
<p>VBA is an object-oriented language. Controls, cells, ranges, shapes, worksheets, and workbooks are all different types of objects that can be manipulated using VBA in Excel. Objects have three types of features: properties, methods, and events.</p>
<h2>Properties</h2>
<p>Properties are attributes associated with an object. For example, a command button has properties for the control name, the caption, the size, the font, and the colors. You can change properties using the Properties window or in code using an assignment statement:</p>

<pre><code class="language-vb">Range("B5").Formula = "=A5/12 + B15"
</code></pre>

<p>Properties can be thought of as a collection of variables associated with an object. Each properties has a type (integer, double, string) and some properties are other objects. For example, the following line of code shows how to modify a property associated with an embedded object.</p>
<pre><code class="language-vb">Range("B5").Interior.Color = 65535
</code></pre>


<h2>Methods</h2>
<p>Methods are actions that can be performed by an object. For example, a range object has a Select method that selects the cells in the range.</p>

<pre><code class="language-vb">Range("B4:D12").Select
</code></pre>

<p>The following Method clears the data in a set of cells while keeping the formatting intact:</p>

<pre><code class="language-vb">Range("B5").ClearContents
</code></pre>

<p>Not all objects include methods.</p>
<h2>Events</h2>
<p>An event is something that happens to an object. We write code in response to events. Events are described in more detail on the <a href="../vb-events/index.php">Working with Events</a> section.</p>
<h2>Collections</h2>
<p>A collection is a special type of object that is a set of objects. For example, the Sheets collection contains all of the Sheet objects in a workbook. For each sheet, there is a collection of Shape objects called Shapes. There is also a collection of Comment objects called Comments. All collection objects contain a Count property that tracks the number of objects in the collection. To reference the objects in a collection, you can use the index of the object:</p>

<pre><code class="language-vb">Sheets(2).Activate</code></pre>

<p>Or if you know the name of the object, you can reference it by name:</p>

<pre><code class="language-vb">Sheets("Sheet2").Activate</code></pre>

<p>To traverse a collection, we use the <strong>For Each</strong> type of loop. For example, to delete all of the shapes on a page, we could do the following:</p>

<pre><code class="language-vb">Dim myshape As Shape
For Each myshape in Shapes
	myshape.Delete
Next myshape</code></pre>

<p>The For Each loop is explained in more detail in the <a href="../vb-loops/">Loops</a> chapter.</p>
<h2>Object Browser</h2>
<p>When we work objects in Excel, we are using the <strong>Microsoft Excel Object Model</strong>. The model can be explored in the Visual Basic Editor using the <strong>Object Browser</strong>. To access the browser, you click on the Object Browser icon <img src="obicon.png" width="20" height="19" />. This brings up the Object Browser window:</p>
<p><img src="browser.png" width="918" height="686" /></p>
<p>The Object Browser is a great way to explore objects and learn what members (properties, methods) are available with the object. The top left item lets us pick which object model to browse. You typically want to change this to &quot;Excel&quot; (as shown) to focus on the Excel objects rather than browsing all objects in MS Office. You can then use the search field (to the left of the binocular icon) to search for a particular object type. You then pick the object of interest in the search results and the members of the object (properties and methods) are listed in the main part of the window. You can click on each member and get a summary at the bottom of the window. You can also right-click on a member to bring up context sensitive help on that member. For example, doing this:</p>
<p><img src="rightclick.png" width="300" height="332" /></p>
<p>brings up this:</p>
<p><img src="help.png" width="797" height="751" /></p>
<p>Of course, you can also look up information on objects directly using the search feature in the Visual Basic Help utility.</p>

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
      <td> <strong>Project Management -</strong> Use a variety of objects and change their appearance and/or values through methods and properties to update the project management table.</td>
      <td align="center" valign="top">Easy</td>
      <td align="center" valign="top"><a href="project_management.xlsm">project_management.xlsm</a></td>
      <td align="center" valign="top"><a href="project_management_key.xlsm">project_management_key.xlsm</a></td>
    </tr>
    <tr>
      <td><strong>Objects - </strong>Run through a few object.method basics to obtain specific results and get rid of an Excel Bug :)</td>
      <td align="center" valign="top">Medium</td>
      <td align="center" valign="top"><a href="objects.xlsm">objects.xlsm</a></td>
      <td align="center" valign="top"><a href="objects_key.xlsm">objects_key.xlsm</a></td>
    </tr>
    <tr>
      <td><strong>Pop the Balloon -</strong> Use shape objects and use methods and properties to create an animation that pumps up a balloon until it pops! </td>
      <td align="center" valign="top">Hard</td>
      <td align="center" valign="top"><a href="pop_the_balloon.xlsm">pop_the_balloon.xlsm</a></td>
      <td align="center" valign="top"><a href="pop_the_balloon_key.xlsm">pop_the_balloon_key.xlsm</a></td>
    </tr>
  </tbody>
</table>
<br />

<?php 
require "../footer.php";
?>

</body>
</html>
