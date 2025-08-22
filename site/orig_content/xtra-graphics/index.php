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

<h1>Custom Graphics</h1>
<p>A common task faced by programmers is how to display custom graphics using source 
  code.&nbsp; It is often useful to display an object that is 
  properly dimensioned in terms of the input parameters supplied on the user.&nbsp; 
  For example, one could display the geometry of a cantilever beam or a column 
  based on the user input.&nbsp; 
  At the other end of the spectrum, it is possible to write sophisticated computer 
  programs with 3D graphics and animation.</p>
<p>Standard VB (applied to a VB form) has a simple, yet powerful set of graphics 
  options.&nbsp; You create a Picture object and then use a series of commands to 
  draw lines and simple shapes in the Picture object.&nbsp; However, none of these 
  tools can used for VBA in Excel.&nbsp; With Excel, an entirely different 
  approach must be used.&nbsp; This approach involves a special type of object 
  called a &quot;Shape&quot;.&nbsp; Shapes can be created manually by the user of the 
  spreadsheet using the standard MS Office drawing tools:</p>
<p><img src="shapespicker.png" width="520" height="641" /></p>
<p>Any of the graphical objects in this menu (lines, connectors, basic shapes, 
  etc.) are classified as shapes.&nbsp; Once created, they can be manipulated via 
  VB code.&nbsp; Since the basic shapes include lines, rectangles, circles, and 
  polygons, you can create just about any custom drawing that you can think of.</p>
<h2>The Shape Object</h2>
<p>When dealing with shapes in VB code, we use the <b>Shape</b> object.&nbsp; 
  All of the objects in the drawing layer of a worksheet, including AutoShapes, 
  freeforms, OLE objects, or pictures, are Shape type objects.</p>
<p>To declare a variable as a Shape object, do the following:</p>

<pre><code class="language-vb">Dim sh As Shape
</code></pre>

<p>The Creating Shapes section below discusses how to create Shape objects.</p>
<h2>The Shapes Collection</h2>
<p>All of the Shape type objects associated with a specific sheet are organized 
  into a set of objects called the <b>Shapes</b> collection.&nbsp; The Shapes 
  collection is a special type of object that has it's own unique set of 
  properties and methods.&nbsp; For example, you can traverse through all of the 
  Shape objects in the Shapes collection using the following code.</p>
  
<pre><code class="language-vb">'Check to see if there is already a polygon named "mypolygon"
'If so, we will delete it.
Dim sh As Shape
For Each sh In Shapes
    If sh.Name = "mypolygon" Then
        sh.Delete
    End If
Next sh
</code></pre>  
  

<p>When you create a new Shape object, it is added to the Shapes collection.</p>
<h2>Creating Shapes</h2>
<p>The simplest way to create new shapes is to use one of the &quot;Addxxx&quot; methods 
  associated with the Shapes collection.&nbsp; These methods include the AddLine, 
  AddPolyline, and AddShape methods.&nbsp; Each of these methods creates a new 
  shape object that is added to the Shapes collection.</p>
<h3>The AddLine Method</h3>
<p>The <b>AddLine</b> method creates a simple line defined by xy coordinates of the 
  beginning and end of the line.&nbsp; The syntax for the method is:</p>
  
<pre><code class="language-vb">expression.AddLine(BeginX, Beginy, EndX, EndY)
</code></pre> 
  
<p>where <i>expression</i> is a Shapes type object.&nbsp; For example, the 
  following code:</p>

<pre><code class="language-vb">Shapes.AddLine 10, 10, 250, 250
</code></pre>  
  
<p>creates a line that starts at the coordinates (10, 10) and ends at (250, 250) 
  and adds it to the Shapes collection for the active worksheet.&nbsp; If you want 
  to be more explicit about which sheet the shape is assigned to, you can use the 
  following code:</p>
  
<pre><code class="language-vb">Worksheets(1).Shapes.AddLine 10, 10, 250, 250
</code></pre>
  
<p>or</p>

<pre><code class="language-vb">Worksheets("Sheet1").Shapes.AddLine 10, 10, 250, 250
</code></pre>


<h3>The AddPolyline Method</h3>
<p>The <b>AddPolyline</b> method creates a sequence of line segments defined by 
  a list of coordinates.&nbsp; If the first coordinate is repeated at the end of 
  the list, the method creates a closed polygon.&nbsp; The syntax for the method 
  is:</p>
  
<pre><code class="language-vb">expression.AddPolyline(SafeArrayOfPoints)
</code></pre>
  

<p>where <i>expression</i> is a Shapes type object and <i>SafeArrayOfPoints</i> is a 2D 
  array of Singles representing the coordinates of the polygon.&nbsp; For example, 
  the following code creates a polygon representing a triangle (from the VBA Excel 
  Help File):</p>
  
<pre><code class="language-vb">Dim triArray(1 To 4, 1 To 2) As Single
triArray(1, 1) = 25    'x coordinate of vertex 1
triArray(1, 2) = 100   'y coordinate of vertex 1
triArray(2, 1) = 100
triArray(2, 2) = 150
triArray(3, 1) = 150
triArray(3, 2) = 50
triArray(4, 1) = 25    ' Last point has same coordinates as first
triArray(4, 2) = 100
Shapes.AddPolyline triArray
</code></pre>  
  

<p>Once again, the object is added to the Shapes collection for the current 
  sheet.</p>
<h3>The AddShape Method</h3>
<p>The AddShape method can be used to create a new Shape object that is an <b> AutoShape</b>.&nbsp; The syntax for the method is:</p>

<pre><code class="language-vb">expression.AddShape(Type, Left, Top, Width, Height)
</code></pre>

<p>where <i>expression</i> is a Shapes collection, <i>Type</i> is the type of 
  AutoShape, and <i>Left, Top, Width</i>, and <i>Height</i> are singles defining 
  the location and size of the object.&nbsp; For example, the following code 
  creates a rectangle:</p>
  
 <pre><code class="language-vb">Shapes.AddShape msoShapeRectangle, 25, 50, 150, 200
</code></pre> 
  
<p>The msoShapeRectangle is a VB constant defining the AutoShape type.&nbsp; The 
  following are all legal AutoShape constants:</p>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse; border-width: 0" id="AutoNumber1" width="842">
  <tr>
    <td style="border-style: none; border-width: medium" width="261"><span style="font-family: Arial; font-weight: 700"><font size="2"> msoShape16pointStar <br />
      msoShape24pointStar <br />
      msoShape32pointStar <br />
      msoShape4pointStar <br />
      msoShape5pointStar <br />
      msoShape8pointStar <br />
      msoShapeActionButtonBackorPrevious <br />
      msoShapeActionButtonBeginning <br />
      msoShapeActionButtonCustom <br />
      msoShapeActionButtonDocument <br />
      msoShapeActionButtonEnd <br />
      msoShapeActionButtonForwardorNext <br />
      msoShapeActionButtonHelp <br />
      msoShapeActionButtonHome <br />
      msoShapeActionButtonInformation <br />
      msoShapeActionButtonMovie <br />
      msoShapeActionButtonReturn <br />
      msoShapeActionButtonSound <br />
      msoShapeArc <br />
      msoShapeBalloon <br />
      msoShapeBentArrow <br />
      msoShapeBentUpArrow <br />
      msoShapeBevel <br />
      msoShapeBlockArc <br />
      msoShapeCan <br />
      msoShapeChevron <br />
      msoShapeCircularArrow <br />
      msoShapeCloudCallout <br />
      msoShapeCross <br />
      msoShapeCube <br />
      msoShapeCurvedDownArrow <br />
      msoShapeCurvedDownRibbon <br />
      msoShapeCurvedLeftArrow <br />
      msoShapeCurvedRightArrow <br />
      msoShapeCurvedUpArrow <br />
      msoShapeCurvedUpRibbon <br />
      msoShapeDiamond <br />
      msoShapeDonut <br />
      msoShapeDoubleBrace <br />
      msoShapeDoubleBracket <br />
      msoShapeDoubleWave <br />
      msoShapeDownArrow <br />
      msoShapeDownArrowCallout <br />
      msoShapeDownRibbon <br />
      msoShapeExplosion1 <br />
      msoShapeExplosion2 </font></span></td>
    <td style="border-style: none; border-width: medium" width="302"><span style="font-family: Arial; font-weight: 700"><font size="2"> msoShapeFlowchartAlternateProcess</font><font size="2"> <br />
      msoShapeFlowchartCard <br />
      msoShapeFlowchartCollate <br />
      msoShapeFlowchartConnector <br />
      msoShapeFlowchartData <br />
      msoShapeFlowchartDecision <br />
      msoShapeFlowchartDelay <br />
      msoShapeFlowchartDirectAccessStorage <br />
      msoShapeFlowchartDisplay <br />
      msoShapeFlowchartDocument <br />
      msoShapeFlowchartExtract <br />
      msoShapeFlowchartInternalStorage <br />
      msoShapeFlowchartMagneticDisk <br />
      msoShapeFlowchartManualInput <br />
      msoShapeFlowchartManualOperation <br />
      msoShapeFlowchartMerge <br />
      msoShapeFlowchartMultidocument <br />
      msoShapeFlowchartOffpageConnector <br />
      msoShapeFlowchartOr <br />
      msoShapeFlowchartPredefinedProcess <br />
      msoShapeFlowchartPreparation <br />
      msoShapeFlowchartProcess <br />
      msoShapeFlowchartPunchedTape <br />
      msoShapeFlowchartSequentialAccessStorage <br />
      msoShapeFlowchartSort <br />
      msoShapeFlowchartStoredData <br />
      msoShapeFlowchartSummingJunction <br />
      msoShapeFlowchartTerminator <br />
      msoShapeFoldedCorner <br />
      msoShapeHeart <br />
      msoShapeHexagon <br />
      msoShapeHorizontalScroll <br />
      msoShapeIsoscelesTriangle <br />
      msoShapeLeftArrow <br />
      msoShapeLeftArrowCallout <br />
      msoShapeLeftBrace <br />
      msoShapeLeftBracket <br />
      msoShapeLeftRightArrow <br />
      msoShapeLeftRightArrowCallout <br />
      msoShapeLeftRightUpArrow <br />
      msoShapeLeftUpArrow <br />
      msoShapeLightningBolt <br />
      msoShapeLineCallout1 <br />
      msoShapeLineCallout1AccentBar <br />
      msoShapeLineCallout1BorderandAccentBar <br />
      msoShapeLineCallout1NoBorder</font></span></td>
    <td style="border-style: none; border-width: medium" width="279"><p class="MsoNormal"><b><font face="Arial" size="2">msoShapeLineCallout2 <br />
      msoShapeLineCallout2AccentBar <br />
      msoShapeLineCallout2BorderandAccentBar <br />
      msoShapeLineCallout2NoBorder <br />
      msoShapeLineCallout3 <br />
      msoShapeLineCallout3AccentBar <br />
      msoShapeLineCallout3BorderandAccentBar <br />
      msoShapeLineCallout3NoBorder <br />
      msoShapeLineCallout4 <br />
      msoShapeLineCallout4AccentBar <br />
      msoShapeLineCallout4BorderandAccentBar <br />
      msoShapeLineCallout4NoBorder <br />
      msoShapeMixed <br />
      msoShapeMoon <br />
      msoShapeNoSymbol <br />
      msoShapeNotchedRightArrow <br />
      msoShapeNotPrimitive <br />
      msoShapeOctagon <br />
      msoShapeOval <br />
      msoShapeOvalCallout <br />
      msoShapeParallelogram <br />
      msoShapePentagon <br />
      msoShapePlaque <br />
      msoShapeQuadArrow <br />
      msoShapeQuadArrowCallout <br />
      msoShapeRectangle <br />
      msoShapeRectangularCallout <br />
      msoShapeRegularPentagon <br />
      msoShapeRightArrow <br />
      msoShapeRightArrowCallout <br />
      msoShapeRightBrace <br />
      msoShapeRightBracket <br />
      msoShapeRightTriangle <br />
      msoShapeRoundedRectangle <br />
      msoShapeRoundedRectangularCallout <br />
      msoShapeSmileyFace <br />
      msoShapeStripedRightArrow <br />
      msoShapeSun <br />
      msoShapeTrapezoid <br />
      msoShapeUpArrow <br />
      msoShapeUpArrowCallout <br />
      msoShapeUpDownArrow <br />
      msoShapeUpDownArrowCallout <br />
      msoShapeUpRibbon <br />
      msoShapeUTurnArrow <br />
      msoShapeVerticalScroll <br />
      msoShapeWave </font></b></p></td>
  </tr>
</table>
<h2>Modifying Properties of Shapes</h2>
<p>Once a shape is created, the properties of the shape can be modified using VB 
  code.&nbsp; For example, the following code creates a polygon and then sets some 
  of the properties of the polygon such as the color and the fill style:</p>
  
<pre><code class="language-vb">Dim sh As Shape

Set sh = Shapes.AddPolyline(triArray)

With sh
    .Name = "mypolygon"
    .Fill.ForeColor.RGB = vbBlue
    .Fill.Solid
End With
</code></pre>
  

<p>Note that we must use the <b>Set</b> command to assign the <b>sh</b> variable 
  to the value returned by the <b>AddPolyline</b> method.&nbsp; This style must be 
  used for all assignment statements involving objects.&nbsp; Another way to 
  achieve the same thing would be as follows:</p>

<pre><code class="language-vb">With Shapes.AddPolyline(triArray)
    .Name = "mypolygon"
    .Fill.ForeColor.RGB = vbBlue
    .Fill.Solid
End With
</code></pre>
  
<p>In other words, we can skip the <b>sh</b> variable and assign the properties 
  at the same time that we create the Shape object.</p>
<p>Notice that the Addxxx methods can be called as either functions or sub 
  procedures.&nbsp; When you call it as a function you should put the arguments in 
  parentheses.&nbsp; When you call it as a sub, you should not use parentheses.&nbsp; 
  For example, the following line calles the AddPolyline method as a sub 
  procedure:</p>

<pre><code class="language-vb">Shapes.AddPolyline triArray
</code></pre>
  

<p>While the following code calls the same method as a function:</p>

<pre><code class="language-vb">Set sh = Shapes.AddPolyline(triArray)
</code></pre>

<h2>Deleting Shapes</h2>
<p>In most cases involving custom graphics, you will have a plot that gets 
  updated each time the user changes the input.&nbsp; Theoretically, when you 
  redraw the plot, you could simply resize the existing shapes or you could create 
  a new set of shapes.&nbsp; Unfortunately, it is difficult or impossible to 
  resize some shapes.&nbsp; Therefore, each time you draw your plot, you will be 
  creating a new set of shape objects.&nbsp; What happens to the old shapes when 
  we create the new shapes?&nbsp; You certainly don't want to create the new 
  shapes on top of the old shapes, or you will get a mess.</p>
<p>My solution to this problem is to give each of my shapes a unique name when I 
  create it (see the sample code for the AddPolyline method above).&nbsp;&nbsp; 
  Then, right before I draw my new shapes, I loop through all of the existing 
  shapes and delete the current instances of my custom shapes.&nbsp; The following 
  code searches through the Shapes collection and deletes two lines and a 
  rectangle:</p>
  
<pre><code class="language-vb">Private Sub Remove_MyShapes()
'Check to see if we have already created our shapes
'If so, we will delete them.
Dim sh As Shape
For Each sh In Shapes
    If sh.Name = "myrect" Or sh.Name = "line1" Or sh.Name = "line2" Then
        sh.Delete
    End If
Next sh
End Sub
</code></pre>  
  

<p>When you use this approach, you have to be sure you name your shapes 
  consistently.</p>
<h2>Coordinate Transformations</h2>
<p>Perhaps the most important (and potentially the most difficult) part of 
  creating shape objects in VB code is to make sure that the objects are created 
  at the proper location and at the proper size on the spreadsheet.&nbsp; Note 
  that all of the methods described above for creating shapes are defined in terms 
  of some coordinate system.&nbsp; The default coordinate system for Excel is 
  defined as follows:</p>
<blockquote>
  <p> <img border="0" src="excel_coord_sys.jpg" width="233" height="197" /></p>
</blockquote>
<p>In other words, the upper left corner of the spreadsheet is the origin (0,0) 
  with x increasing to the right and y increasing to the bottom.&nbsp; The 
  coordinates are based on pixels.</p>
<p>The default coordinate system is not always very helpful.&nbsp; Typically, we 
  want to define the coordinates of the Shapes using our own custom coordinate 
  system using a tradition orientation (y is positive in the up direction).&nbsp; 
  In order to do this, we must set up a coordinate  transformation 
  between our custom coordinate system which we call the <b>world</b> coordinate 
  system and the screen coordinate system.&nbsp; The math that is used to perform 
  this coordinate transformation is fairly simple, but I won't describe it in 
  detail here.&nbsp; You can consult any book on basic computer graphics for a 
  full explanation.&nbsp; Rather, I will focus on how to set up and use the 
  transformation.</p>
<p>The first step in setting up the transformation is to define a set of four 
  transformation variables as follows:</p>
  
<pre><code class="language-vb">'Coefficients for coordinate transformation
Dim a As Double
Dim b As Double
Dim c As Double
Dim d As Double
</code></pre>  
  

<p>These variables can be local or global.&nbsp; The following source code 
  examples all assume that 
  they have been defined as global variables at the top of your source code.</p>
<p>Once these variables are set up properly (which will be discussed below), we 
  can transform world coordinates to screen coordinates using the following code:</p>
  
<pre><code class="language-vb">Private Sub transform_coords(ByRef x As Double, ByRef y As Double)

    x = x * a + b
    y = y * c + d

End Sub
</code></pre>
  

<p>For example, the following code creates a <b>Line</b> shape with the world 
  coordinates <b>startxy = (20,20)</b> and <b>endxy = (100,120)</b>:</p>
  
<pre><code class="language-vb">Dim x1 As Double
Dim y1 As Double
Dim x2 As Double
Dim y2 As Double

x1 = 20
y1 = 20
x2 = 100
y2 = 120

transform_coords x1, y1
transform_coords x2, y2

Shapes.AddLine x1, y1, x2, y2
</code></pre> 
  

<p>Before calling the <b>transform_coords</b> sub, we must first initialize the values of the transformation variables.&nbsp; 
  When we do this, we went to set up the transformation so that any world 
  coordinates are transformed into a specific place on our spreadsheet.&nbsp; It 
  is pretty easy to decide on a range in our spreadsheet where we want something 
  to be drawn, but how do we determine the screen coordinates for that range?&nbsp; 
  I have found that the best way to do this is to identify the range as a Range 
  object in VB code and use the .Left, .Right, .Width, and .Height properties of 
  the range.</p>
<p>The following VB sub takes the dimensions of world coordinate range to be 
  used for the graphics and a Range object defining the screen location of the 
  graphics and it initializes the value of the transformation coordinates such 
  that the objects defined in the specified world coordinate range will be 
  centered in the screen range with at least a 10% cushion on the sides (left, 
  right, top, bottom).&nbsp; The aspect ratio is always preserved.&nbsp; This 
  means that the cushion in the vertical direction (or the horizontal direction) 
  may end up being larger than 10%.</p>
  
<pre><code class="language-vb">Private Sub set_up_transformation(xmin As Double, _
                                  xmax As Double, _
                                  ymin As Double, _
                                  ymax As Double, _
                                  drawrange As Range)
'Sets up the global transformation matrix so that any coordinates within the given
'world coordinate bounds will be drawn inside of the specified range with a 20%
'cushion.

Dim xdomain As Double
Dim ydomain As Double
Dim plotratio As Double
Dim polyratio As Double
Dim plotwidth As Double
Dim plotheight As Double
Dim plotleft As Double
Dim plotbot As Double
Dim Sx As Double
Dim Sy As Double

xdomain = xmax - xmin
ydomain = ymax - ymin

polyratio = ydomain / xdomain
plotratio = drawrange.Height / drawrange.Width

If polyratio > plotratio Then
    'y range is the dominant range
    plotwidth = ydomain / plotratio * 1.2
    plotheight = ydomain * 1.2
Else
    'y range is the dominant range
    plotwidth = xdomain * 1.2
    plotheight = xdomain * plotratio * 1.2
End If

plotleft = (xmax + xmin) / 2 - plotwidth / 2
plotbot = (ymin + ymax) / 2 - plotheight / 2

Sx = drawrange.Width / plotwidth
Sy = -drawrange.Height / plotheight

a = Sx
b = drawrange.Left - Sx * plotleft
c = Sy
d = (drawrange.Top + drawrange.Height) - Sy * (plotbot)

End Sub
</code></pre>
  

<p>Here is a sample call to this sub that sets up the mapping over a world 
  coordinate x range of 10 to 100 and a y range of -20 to 120:</p>

<pre><code class="language-vb">set_up_transformation 10, 100, -20, 120, Range("F12:L20")
</code></pre>

<p>In summary, the following code initializes the transformation variables, maps 
  the coordinates, and draws a line:</p>
  
<pre><code class="language-vb">Dim x1 As Double
Dim y1 As Double
Dim x2 As Double
Dim y2 As Double

set_up_transformation 20, 100, 20, 120, Range("F12:L20")

x1 = 20
y1 = 20
x2 = 100
y2 = 120

transform_coords x1, y1
transform_coords x2, y2

Shapes.AddLine x1, y1, x2, y2
</code></pre>  
  

<p>Note that this code only draws one line.&nbsp; If you want to draw 
  multiple objects in the same window using the same coordinate mapping, you need 
  to be sure to call the <code>set_up_transformation </code>sub once before you 
  transform any coordinates and make sure that xmin, xmax, ymin, &amp; ymax are set up 
  according to the limits of <b>all</b> of the objects combined.</p>
<?php 
require "../footer.php";
?>
</body>
</html>
