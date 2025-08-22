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

<h1> Creating Charts</h1>

<p>One of the most fundamental features of Excel is the Charts tool. Charts are used to generate a graphical representation of a set of data. Charts can be incredibly powerful in illustrating trends and characteristics of a data set. In this chapter, we will cover a brief overview of the chart tools with a special emphasis on the types of charts most commonly used in engineering and scientific applications. It is not intended to be an comprehensive overiew of all of the chart options. Such an overview would be beyond the scope of this Primer.</p>
<h2>Chart Types</h2>
<p>The first step in creating a chart is selecting the type of chart to use. This will depend primarily on the type of data that you wish to graph with the chart. The following table lists the more commonly-used charts and the suggested applications:</p>
<table width="569" border="1">
  <tr>
    <th width="50" align="center" scope="col">Type</th>
    <th width="100" align="center" scope="col">Name</th>
    <th width="405" scope="col">Description</th>
  </tr>
  <tr>
    <td align="center"><img src="type-column.png" width="16" height="16" alt=""/></td>
    <td align="center">Column</td>
    <td>Use this chart to visually compare values across a few categories.</td>
  </tr>
  <tr>
    <td align="center"><img src="type-bar.png" width="16" height="16" alt=""/></td>
    <td align="center">Bar</td>
    <td>Use this chart to visually compare values across a few categories when the chart shows duration or the category text is long.</td>
  </tr>
  <tr>
    <td align="center"><img src="type-line.png" width="16" height="16" alt=""/></td>
    <td align="center">Line</td>
    <td>Use this chart to show trends over time (years, months, and days) or categories.</td>
  </tr>
  <tr>
    <td align="center"><img src="type-area.png" width="16" height="16" alt=""/></td>
    <td align="center">Area</td>
    <td>Use this chart to show trends over time (years, months, and days) or categories. Use it to hightlight the magnitude of change over time.</td>
  </tr>
  <tr>
    <td align="center"><img src="type-pie.png" width="16" height="16" alt=""/></td>
    <td align="center">Pie</td>
    <td>Use this chart to show proportions of a whole. Use it when the total of your numbers is 100%.</td>
  </tr>
  <tr>
    <td align="center"><img src="type-scatter.png" width="16" height="16" alt=""/></td>
    <td align="center">Scatter (X,Y)</td>
    <td>Use this chart type to show the relationship between sets of values.</td>
  </tr>
</table>
<p>For scientific and engineering applications, the most common type of chart is the <strong>Scatter (X,Y)</strong> chart, which is sometimes called an <strong>XY Scatter</strong> chart. As the &quot;XY&quot; part of the name implies, this chart is used to represent one set of data (Y) which is dependent upon, or related to another set of data (X), both of which are numeric values. In other words:</p>
<blockquote>
  <p>y = f(x)</p>
</blockquote>
<p>or y is some function of x. This can be an explicit numerical function (y = x<sup>2</sup>-3x+1) or it could be an implicit relationship, such as measured strength of some specimens as a function of applied load.</p>
<h2>Creating a Chart</h2>
<p>The steps to creating a new chart are as follows:</p>
<ol>
  <li>Select the data in the sheet that will be associated with the chart.</li>
  <li>Select the <strong>Insert</strong> tab.</li>
  <li>In the <strong>Charts</strong> section, click on one of the chart type icons and then select the specific type of chart you wish to create.</li>
</ol>
<p>To illustrate the process, consider the following example worksheet. This is a variation of the parabola worksheet described in the <a href="../excel-goalseek/">Goal Seek and Solver</a> chapter.</p>
<p><img src="parabola-start.png" width="425" height="499" alt=""/></p>
<p>Our objective is to create an XY Scatter chart of the XY values shown in the tables. These values represent a solution of the equation:</p>
<blockquote>
  <p>y = x<sup>2</sup> - 3x + 1</p>
</blockquote>
<p>for a the range of x values varying from -1 to 4. To create the chart, we select the cells in the range <strong>B12:B22</strong> and follow the steps outlined above as follows:</p>
<p><img src="insert-chart.png" width="893" height="697" alt=""/></p>
<p>Note that the chart type selected was <strong>Scatter with Smoth Lines</strong>. The &quot;Smooth Lines&quot; part means that a smooth curve is fit the to XY points that interpolates the points and provides a natural curvature between the points using some type of spline function. This is typically the best option to select. By contrast, this is what the &quot;Straight Lines&quot; option looks like:</p>
<p><img src="straight-lines.png" width="486" height="296" alt=""/></p>
<p>Markers can also be combined  with the smooth or straight lines. A marker is a dot at the location of each XY coordinate pair. Here is the <strong>Markers with Smooth Lines</strong> option:</p>
<p><img src="markers-lines.png" width="486" height="294" alt=""/></p>
<p>And the <strong>Markers Only</strong> option:</p>
<p><img src="markers-only.png" width="486" height="294" alt=""/></p>
<p>As a matter of style, markers should only be used when there is a some kind of significance to each of the XY pairs. For example, perhaps the XY pairs represent data collected in the field or lab and each point corresponds to a sample or measurement. In many cases, however, the XY values represent some underlying fuction (such as the case shown above) and the points are abritrarily selected. In this case, <strong>markers should not be used</strong> as they simply detract from the display of the function.</p>
<h2>Formatting a Chart</h2>
<p>Once the chart is created, we can edit the chart options to modify the formatting. If you click on a chart, a set of three buttons will appear just to the right of the charte. The &quot;+&quot; button can be used to add or remove chart elements such as axis labels, the chart title, and a legend.</p>
<p><img src="chart-elements.png" width="708" height="317" alt=""/></p>
<p>After editing the chart title and axis labels, the chart looks like this:</p>
<p><img src="chart-before-axis-edit.png" width="486" height="296" alt=""/></p>
<p>Note that the range on the x- and y-axes are automatically determined. Suppose for this case that we wish to limit the range of the x-axis to vary from -1 to 4. To do this we double-click on the x-axis or right-click on the axis and select Format Axis. This brings up the Format Axis options on the right side:</p>
<p><img src="format-axis-1.png" width="275" height="337" alt=""/></p>
<p>To remove the &quot;Auto&quot; option for the max and min bounds, we simply type in new values and hit the Enter key. After doing so, the Axis Options display as follows:</p>
<p><img src="format-axis-2.png" width="272" height="318" alt=""/></p>
<p>Clicking the Reset button would revert back to the automatic setting. After manually editing the x-axis bounds, the chart looks like this:</p>
<p><img src="chart-after-axis-edit.png" width="490" height="296" alt=""/></p>
<h2>Changing the Data Source</h2>
<p>In some cases, after creating the chart we wish to change the set of cells associated with the chart (i.e., the &quot;data source&quot;). For example, perhaps we have deleted some of our XY pairs or we have extended the table to add additional pairs. When we do so, the chart is not automatically updated to reflect the change; we must manually make the correction. To change the data source, you can do  the following:</p>
<ol>
  <li>Click on the curve in the chart to select it. This will display the range of cells associated with the chart.</li>
  <li>Using the handles at the corners of the highlighted ranges, drag the corners to resize the selection to the desired range.</li>
</ol>
<p><img src="data-source-1.png" width="816" height="336" alt=""/></p>
<p>Another option for changing the source is:</p>
<ol>
  <li>Right-click on the curve and select the <strong>Select Data</strong> option. This brings up the <strong>Select Data Source</strong> dialog. </li>
  <li>Edit the <strong>Chart data range</strong> field to correspond to the desired range.</li>
</ol>
<p><img src="data-source-2.png" width="598" height="326" alt=""/></p>
<h2>Sample Workbook</h2>
<p>The workbook used in the examples shown above can be downloaded here:</p>
<p><a href="parabola2.xlsx">parabola2.xlsx</a></p>

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
      <td> <strong>Excess Pore Pressure -</strong> Create a chart of the excess pore pressure vs distance from given test data. </td>
      <td align="center" valign="top">Easy</td>
      <td align="center" valign="top"><a href="excess_pore_pressure.xlsx">excess_pore_pressure.xlsx</a></td>
      <td align="center" valign="top"><a href="excess_pore_pressure_key.xlsx">excess_pore_pressure_key.xlsx</a></td>
    </tr>
    <tr>
      <td><strong>Material Price Trends - </strong>Create a chart of the prices for different engineering materials over a specified date range and identify a trend.</td>
      <td align="center" valign="top">Medium</td>
      <td align="center" valign="top"><a href="material_price_trends.xlsx">material_price_trends.xlsx</a></td>
      <td align="center" valign="top"><a href="material_price_trends_key.xlsx">material_price_trends_key.xlsx</a></td>
    </tr>
    <tr>
      <td><strong>Crater Settlement-</strong> Plot the settlement vs distance of different points due to underground blasting. Analyze the crater formed. </td>
      <td align="center" valign="top">Hard</td>
      <td align="center" valign="top"><a href="crater_settlement.xlsx">crater_settlement.xlsx</a></td>
      <td align="center" valign="top"><a href="crater_settlement_key.xlsx">crater_settlement_key.xlsx</a></td>
    </tr>
  </tbody>
</table>
<br />

<?php 
require "../footer.php";
?>
</body>
</html>
