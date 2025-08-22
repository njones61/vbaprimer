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

<h1> Annotation - Drawings and Equations</h1>

<p>When developing worksheets as a tool for solving engineering problems that we frequently encounter in practice, it is helpful to make the worksheets as clear and easy to use as possible. One of the ways this can be accomplished is to add graphics to the worksheet illustrating the problem being solved. In some cases, showing the equations in native form is also helpful. For example, suppose we are just starting to develop a worksheet for solving the deflection of a simply-supported beam subjected to a point load. </p>
<p><img src="beamsheet-start.png" width="715" height="344" alt=""/></p>
<p>The equations associated with the problem are as follows:</p>
<p><img src="beam.png" width="443" height="239" alt=""/></p>
<p><img src="beam_eq.gif" width="447" height="170" alt=""/></p>
<p>Before entering the formulas, it would be very helpful to future users of the workbook if we could add some annotation to the workbook including a diagram and set of equations similar to those shown above. In this chapter, we will use this example to review the basic tools provided in Excel for making simple drawings and for composing graphical equations.</p>
<p>(Note: You may want to open the sample workbook at the bottom of this chapter and follow along.)</p>
<h2>Drawing Tools</h2>
<p>The drawing tools in Excel are the same set of tools utilized in other Microsoft Office products, including Word and Powerpoint. Drawings are created using a set of objects called <strong>shapes</strong>. You begin creating a drawing in Excel by inserting a shape. To insert a shape, click on the <strong>Insert</strong> tab, and then click on the <strong>Shapes</strong> menu in the <strong>Illustrations</strong> section.</p>
<p><img src="shapes.png" width="662" height="698" alt=""/></p>
<p>To create our drawing, we will first create the lines corresponding to the x and y axes. Since the have an arrow at the ends, we will use the line shape with an arrow on one end (<img src="shape-arrow.png" width="16" height="16" alt=""/>). To create the x-axis, start on the left and click and drag to the right. If you hold the shift key down while dragging, the line will stay perfectly flat. After creating the line, you will shift into drawing mode and the drawing tools will be displayed:</p>
<p><img src="drawingtools.png" width="888" height="464" alt=""/></p>
<p>If you click anywhere on the worksheet so that no shapes are selected, the drawing tool ribbon will go away. To get it back, you can simply click on one of your shapes or insert another shape using the process described above.</p>
<p>Now we need to create the y-axis. Reselect the line/arrow tool from the shape pallette on the left side of the drawing tools ribbon. Then start at the left side of the axis (the origin) and repeat the process dragging straight up, again holding down the shift key while dragging. At this point, the drawing should look like this:</p>
<p><img src="drawing-axes.png" width="454" height="168" alt=""/></p>
<p>Next, we will create the deflected beam shape. The best way to do this is with the curve shape (<img src="shape-curve.png" width="18" height="15" alt=""/>). Select the curve shape. Curves are created by clicking on a set of pivot points defining the curve and then double-clicking at the end. To create the beam, we will click on three points: one on the left side of the beam to start, one at the middle of the beam, then we will double-click on the right side to end. Star creating your beam by clicking on the coordinate origin. Then click at the midpoint of the beam as follows:</p>
<p><img src="curve-1.png" width="445" height="203" alt=""/></p>
<p>The double-click on the right side of the beam:</p>
<p><img src="curve-2.png" width="426" height="184" alt=""/></p>
<p>Note that you have created a nice smooth curve based on these three points. If you want to fine-tune the shape of the curve, select the curve and then select the <strong>Edit Shape|Edit Points</strong> command from the <strong>Insert Shapes</strong> section of the <strong>Drawing Tools</strong> ribbon.</p>
<p><img src="editpoints1.png" width="327" height="165" alt=""/></p>
<p>At this point, the three points defining the curve will appear. If click on one of the points, a set of handles will appear:</p>
<p><img src="editpoints2.png" width="466" height="105" alt=""/></p>
<p>Repositioning the point in the center or dragging the handles will reshape the curve. Experimenting for a few minutes will give you a good sense of how this works. When you are done try to restore the curve to the shape shown above as best you can (or delete and recreate the curve).</p>
<p>Next we will create the support points. In the drawing shown above, the supports are shown with a triangle and circle. To keep things simple we will use a triangle only. Select the triangle shape (<img src="shape-triangle.png" width="15" height="15" alt=""/>) and drag a triangle at the left end of the beam. You may need to create the shape and the drag it into position. You may also need to use the handles at the corners of the shape to resize it. If you want to give the shape a professional look, you can apply one of the &quot;3D&quot; shape styles:</p>
<p><img src="shape-styles.png" width="506" height="382" alt=""/></p>
<p>At this point, the support should look like this:</p>
<p><img src="support-left.png" width="163" height="78" alt=""/></p>
<p>To create a support on the other end, we can either repeat the entire process, or we can create a duplicate of the shape by pressing Ctrl-D  and reposition it to the right end as follows:</p>
<p><img src="support-right.png" width="468" height="80" alt=""/></p>
<p>Next we will create the shape representing the load. To do this we will use the block arrow tool corresponding to a down arrow (<img src="shape-blockarrow.png" width="17" height="17" alt=""/>). Click and draw in a long vertical shape to create the load arrow. Apply one of the 3D styles and the reposition the arrow as follows:</p>
<p><img src="load-arrow.png" width="415" height="174" alt=""/></p>
<p>Now we are ready to apply the lines associated with the dimensioning. Using a combination of the line tools (<img src="shape-arrow.png" width="16" height="16" alt=""/> <img src="shape-arrow2.png" width="16" height="16" alt=""/> <img src="shape-line.png" width="16" height="16" alt=""/>) and the techniques described above, create the following lines:</p>
<p><img src="dimension-lines.png" width="423" height="236" alt=""/><br />
As you create the lines, you may want to simply duplicate the existing lines. Also, if you need to align some of the shapes, select multiple shapes using the shift key and then select the <strong>Align</strong> menu in the <strong>Arrange</strong> section of the <strong>Drawing Tools</strong> ribbon. </p>
<p><img src="aligntools.png" width="355" height="332" alt=""/></p>
<p>The last step is to add the text labels to the drawing. We will start with the &quot;P&quot; by the load arrow. To add the label, select the text box shape (<img src="shape-text.png" width="17" height="15" alt=""/>) and drag a box by the load arrow. Then type &quot;P&quot; and click on the box containing the label to reposition it if necessary. To get right of the line around the text box, select the label and then select the <strong>Shape Outline</strong> menu in the <strong>Shape Styles</strong> section of the ribbon and select <strong>No Outline</strong>. To change the font attributes (size, font, bold, alignment, etc.) you will need to click on the <strong>Home</strong> ribbon and use the tools in the <strong>Font</strong> section while the textbox shape is selected.</p>
<p><img src="text-line-format.png" width="452" height="496" alt=""/></p>
<p>To create the remaining labels, duplicate (Ctrl-D) the current label and then change the text. Reposition and re-format as needed. When you are finished, our drawing is complete!</p>
<p><img src="completedrawing.png" width="411" height="232" alt=""/></p>
<h2>The Equation Editor</h2>
<p>Next we will use the Excel Equation Editor to create a graphical representation of the equations describing the deflection of a simply supported beam. To make things interesting, we will start with the second equation shown here:</p>
<blockquote>
  <p><img src="equation-start.png" width="303" height="55" alt=""/></p>
</blockquote>
<p>To insert an a new equation, select a cell somewhere in the vicinity of where you would like to put the equation and then click on the Insert tab. Over on the far right end of the tab, you will see a Symbols section. Click on the Equation menu.</p>
<p><img src="equation-menu.png" width="666" height="412" alt=""/></p>
<p>If you click on the down arrow, the menu lists some frequently-used equations for cases where you want to start with something and then modify. But we need to start from scratch, so we need to click on top part of the <strong>Equation</strong> button. This inserts a new equation and opends up the <strong>Equation Tools</strong> ribbon.</p>
<p><img src="equation-ribbon.png" width="1324" height="150" alt=""/></p>
<p>The ribbon is divided into two main sections: <strong>Symbols</strong> and <strong>Structures</strong>. The structures represent basic components of equations that define positional relationships. In general, you form an equation by inserting a structure and then filling in the elements of the structure by typing standard characters (x, y, p, t, etc.) or by selecting symbols. We will start our equation by typing &quot;v=&quot;. Then we select the <strong>Fraction</strong> menu and select the first (upper left) option (&quot;stacked fraction&quot;). This creates an empty fraction with a placeholder for the numerator and the denominator:</p>
<blockquote>
  <p><img src="equation-edit-1.png" width="70" height="71" alt=""/></p>
</blockquote>
<p>Click on the numerator box to and type &quot;-Pb&quot;. Then click on the denominator box and type &quot;6E&quot;. At this point, we need to enter &quot;Iu&quot; with the &quot;u&quot; part as a subscript. To do so, we click on the Script menu in the Structures section and select the second item on the top row. This creates a place holder for the main part and the subscript. Click on the main part and type &quot;I&quot; and then click on the subscript and type &quot;u&quot;. (Note hhat you rather than clicking on the place holders, you can use the arrow keys on your keyboard to navigate between the different placeholders.) At this point, the equation should look like this:</p>
<blockquote>
  <p><img src="equation-edit-2.png" width="80" height="72" alt=""/></p>
</blockquote>
<p>At this point, we are finished with the fraction and need to enter the remainder of the equation. Before doing so, we need to move the cursor out of the denominator, otherwise whatever we type will be in the denominator. To do so, we simply hit the right-arrow key on the keyboard a couple of times until the cursor is to right of and even with the main part of the equation:</p>
<blockquote>
  <p><img src="equation-edit-3.png" width="79" height="73" alt=""/></p>
</blockquote>
<p>Next we need to create a set of square brackets for the right side of the equation. We could simply type the &quot;[&quot; and &quot;]&quot; characters on the keyboard, but we need the brackets to automatically resize themselves based on the content in the bracksts, so we need to use a structure. Click on the Brackets menu and select the square brackets option (second item in the top row).</p>
<blockquote>
  <p><img src="equation-edit-4.png" width="105" height="75" alt=""/></p>
</blockquote>
<p>Click on the placeholder in the middle of the brackets and insert another fraction structure. Enter &quot;L&quot; for the numerator and &quot;b&quot; for the denominator.</p>
<blockquote>
  <p><img src="equation-edit-5.png" width="97" height="72" alt=""/></p>
</blockquote>
<p>Then hit the right arrow key once to exit the denominator. </p>
<p>Next, we need to enter (x-a)<sup>3</sup>. To do this, we will need two structures: a superscript and a bracket. First insert the superscript structure, then click on the main part of the superscript structure and enter a bracket &quot;()&quot; structure. Then enter &quot;x-a&quot; in the main part of the bracket structure and click on the superscript and type &quot;3&quot;.</p>
<blockquote>
  <p><img src="equation-edit-6.png" width="161" height="73" alt=""/></p>
</blockquote>
<p>Hit the right arrow key to exit the superscript and finish the rest of the equation using the procedures described above. When you are finished, the equation should look like this:</p>
<blockquote>
  <p><img src="equation-edit-7.png" width="259" height="43" alt=""/></p>
</blockquote>
<p>After creating the equation, you can move it around to reposition it. To edit it, simply click on it and the Equation Tools ribbon will appear again.</p>
<p>You can now continue and try creating the other equations shown above if you wish.</p>
<h2>Selecting Objects</h2>
<p>After creating a set of shapes and equations, it is sometimes necessary to select them as a group by dragging a box around them. For example, this is necessary if you wish to reposition all of them at once. If you just drag a box around the objects, they are not selected. To do this, you need to activate the &quot;Select Objects&quot; mode. This is accomplished by clicking on the <strong>Home</strong> tab and then selecting the <strong>Select Objects</strong> item in the <strong>Find and Select</strong> menu on the far right end of the ribbon.</p>
<p><img src="selectobjects.png" width="378" height="357" alt=""/></p>
<h2>Grouping</h2>
<p>When dealing with large sets of shapes and equations as described in the previous section, it is sometimes helpful to organize the objects into groups. This allows you to select and move them as a single item. To group a set of objects, select them and then select the <strong>Group</strong> item in <strong>Arrange</strong> section in the <strong>Drawing Tools </strong>ribbon. Select the Group option from the menu.</p>
<h2><img src="group.png" width="250" height="182" alt=""/></h2>
<p>The ungroup command can be used to undo the process of grouping a set of objects.</p>
<h2>Display Order</h2>
<p>When working with shapes that are filled, sometimes one objects can obscure another. You can directly control the order in which the objects are displayed using the options in the <strong>Bring Forward</strong> and <strong>Send Backward</strong> menus in the <strong>Arrange</strong> section of the <strong>Drawing Tools</strong> ribbon.</p>
<p><img src="arrange.png" width="298" height="92" alt=""/></p>
<h2>Sample Workbooks</h2>
<p>The workbook used in the examples shown above can be downloaded here:</p>
<table width="926" border="0">
  <tr>
    <td width="153"><a href="simplebeam.xlsx">simplebeam.xlsx</a><br />    </td>
    <td width="763">&lt;-- Without annotation. Start here if you want to try creating the drawings and equations<br /></td>
  </tr>
  <tr>
    <td><a href="simplebeam-key.xlsx">simplebeam-key.xlsx</a></td>
    <td>&lt;-- With completed annotation. </td>
  </tr>
</table>
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
      <td> <strong>Drag Equation -</strong> Create an annotated figure of   drag force on a baseball  and compute the drag force that the ball experiences.</td>
      <td align="center" valign="top">Easy</td>
      <td align="center" valign="top"><a href="drag_equation.xlsx">drag_equation.xlsx</a></td>
      <td align="center" valign="top"><a href="drag_equation_key.xlsx">drag_equation_key.xlsx</a></td>
    </tr>
    <tr>
      <td><strong>Unit Weight - </strong>Create an annotated soil profile and write an equation to compute the unit weight of the two soil layers.</td>
      <td align="center" valign="top">Medium</td>
      <td align="center" valign="top"><a href="unit_weight.xlsx">unit_weight.xlsx</a></td>
      <td align="center" valign="top"><a href="unit_weight_key.xlsx">unit_weight_key.xlsx</a></td>
    </tr>
    <tr>
      <td><strong>Headloss -</strong> Annotate a figure of a pipe with the necessary equations to calculate headloss. Then calculate the headloss with the given values and equations.</td>
      <td align="center" valign="top">Hard</td>
      <td align="center" valign="top"><a href="headloss.xlsx">headloss.xlsx</a></td>
      <td align="center" valign="top"><a href="headloss_key.xlsx">headloss_key.xlsx</a></td>
    </tr>
  </tbody>
</table>
<br />

<?php 
require "../footer.php";
?>
</body>
</html>
