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

<h1> Using the Goal Seek and Solver Tools</h1>

<p>There are many cases when performing computations in Excel where we need to solve an equation that is eitther difficult or impossible to solve directly. Therefore we need to solve it using some sort of iterative process. The <strong>Goal Seek</strong> and <strong>Solver</strong> tools are perfectly suited for these cases. For example, consider the following workbook:</p>
<p><img src="workbook-1.png" width="611" height="489" alt=""/></p>
<p>The workbook is designed to solve a quadratic equation of the form:</p>
<blockquote>
  <p>y = ax<sup>2</sup> + bx + c</p>
</blockquote>
<p>The user enters the coefficients a, b, &amp; c in cells C4:C6. For the case shown above, we are solving:</p>
<blockquote>
  <p>y = x<sup>2</sup> - 3x + 1</p>
</blockquote>
<p>The chart at the bottom is used to graph the parabola corresponding to the equation over a specified range. The solution to the equation is the two points where the parabola intercepts the y axis. These are called the &quot;roots&quot; and represent the solution to:</p>
<blockquote>
  <p>ax<sup>2</sup> + bx + c = 0</p>
</blockquote>
<p>As can be seen on the chart, the roots are approximately 0.4 and 2.6. </p>
<p>We can find the roots using cells F4 and F5. We enter a value for x in cell F4. The corresponding value of y is computed in cell F5 as:</p>
<blockquote>
  <p>=C4*F4^2+C5*F4+C6</p>
</blockquote>
<p>To find a root, we can enter a guess (0.4 for starters) into F4 and iteratively tweak that number until the value computed for y in F5 is roughly equal to zero. While this works, it can be time consuming and tedious.</p>
<h2>Goal Seek</h2>
<p>A more efficient way to solve for the roots is to let Excel perform the iterative calculations using the Goal Seek tool. This tool is located in the <strong>Data</strong> ribbon under the <strong>What-If Analysis</strong> menu. It has three inputs. For the case shown above, the inputs should be like this:</p>
<p><img src="goalseek-1.png" width="226" height="152" alt=""/></p>
<p>This tells the Goal Seek tool to repeatedly change cell F4 (x) until the value in cell F5 (y) is equal to zero. After clicking the OK button, the following message appears:</p>
<p><img src="goalseek-2.png" width="273" height="155" alt=""/></p>
<p>Note that the solution is not always found exactly due to roundoff error. After running the tool, the following values are displayed:</p>
<p><img src="goalseek-3.png" width="426" height="167" alt=""/></p>
<p>To find other root, we need to repeat the process, but we must first enter a value for x that is close to the second root. If we enter an x value of 2.5 and repeat the process, we can quickly find the other root:</p>
<p><img src="goalseek-4.png" width="412" height="159" alt=""/></p>
<h2>Solver</h2>
<p>There is another tool in Excel for performing iterative calculations called the Solver that is even more powerful. The Solver is an Add-In so before using it, there are a few steps we need to take.</p>
<ol>
  <ol>
    <li>Select the <b>File|Open</b> command, and then select <strong> Options</strong>.</li>
    <li>Click <strong>Add-Ins</strong>, and then in the <strong>Manage</strong> box, select <strong>Excel Add-ins</strong>.</li>
    <li>Click <strong>Go</strong>. </li>
    <li>In the <strong>Add-Ins</strong> available box, select the <strong>Solver Add-in</strong> check box, and then click <strong>OK</strong>.</li>
  </ol>
  <blockquote>
    <p>If  <em>Solver Add-in</em> is not listed in the <strong>Add-Ins available</strong> box, click <strong>Browse</strong> to locate the add-in<br />
      If you get prompted that the Solver Add-in is not currently installed on your computer, click <strong>Yes</strong> to install it.</p>
  </blockquote>
</ol>
<p>  After you load the Solver Add-in, the <strong>Solver</strong> command is available in the <strong>Analysis</strong> group on the <strong>Data</strong> tab. These steps only need to be completed once.</p>
<p>After launching the Solver, the following window appears:</p>
<p><img src="solver-1.png" width="584" height="591" alt=""/></p>
<p>In general, the Solver is like Goal Seek in that it iteratively changes one (or more) input cell(s) until some condition is met. But in this case there are three possible conditions (max, min, value of) and a set of constraints can be defined. When we use the <strong>Value Of</strong> option, it is essentially the same as Goal Seek. Using the options shown above, we can solve for one of the roots of the parabola by clicking the <strong>Solve</strong> button. Doing so brings up the following message:</p>
<p><img src="solver-2.png" width="474" height="363" alt=""/></p>
<p>Generally you want to select the OK option to keep the solver solution. The solution found by the solver is:</p>
<p><img src="solver-3.png" width="427" height="164" alt=""/></p>
<p>which is the same as above (the starting value was near the second root), but a little more accurate.</p>
<p>The real power of the solver is to perform optimization using the <strong>Max</strong> and <strong>Min</strong> options. This is something that cannot be done with Goal Seek. For example, suppose we wanted to find the x location corresponding the lowest point on the parabola. We could simply enter a guess for the x value and run the Solver with the following options:</p>
<p><img src="solver-4.png" width="584" height="591" alt=""/></p>
<p>This results in the following solution:</p>
<p><img src="solver-5.png" width="476" height="175" alt=""/></p>
<p>Which is precisely the correct result. </p>
<p>While not applicable in this case, we can enter a series of constraints such as &quot;B4&gt;=0&quot;. As the Solver iterates, a variety of input values are tested. Such constraints can ensure that the Solver algorithm stays stable and will be more likely to converge on a solution.</p>
<h2>Sample Workbook</h2>
<p>The workbook used in the examples shown above can be downloaded here:</p>
<p><a href="parabola.xlsx">parabola.xlsx</a></p>

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
      <td> <strong>Going Fishing -</strong> Use goalseek on a parabolic function to find where the fish strikes the fly.</td>
      <td align="center" valign="top">Easy</td>
      <td align="center" valign="top"><a href="going_fishing.xlsm">going_fishing.xlsm</a></td>
      <td align="center" valign="top"><a href="going_fishing_key.xlsm">going_fishing_key.xlsm</a></td>
    </tr>
    <tr>
      <td><strong>Polynomials - </strong>Use goalseek to find the zeros of a polynomial function.</td>
      <td align="center" valign="top">Medium</td>
      <td align="center" valign="top"><a href="polynomials.xlsm">polynomials.xlsm</a></td>
      <td align="center" valign="top"><a href="polynomials_key.xlsm">polynomials_key.xlsm</a></td>
    </tr>
    <tr>
      <td><strong>Topo Solver - </strong>Use goalseek and solver to identify key points on a topographic map.</td>
      <td align="center" valign="top">Medium</td>
      <td align="center" valign="top"><a href="topo_solver.xlsm">topo_solver.xlsm</a></td>
      <td align="center" valign="top"><a href="topo_solver_key.xlsm">topo_solver_key.xlsm</a></td>
    </tr>
    <tr>
      <td><strong>Missile Trajectory -</strong> A asteroid is headed for earth. You need to determine if the trajectory of a "asteroid-stopping" defense missile will clear the nearby buildings and destroy the asteroid before it impacts with earth.</td>
      <td align="center" valign="top">Hard</td>
      <td align="center" valign="top"><a href="missile_trajectory.xlsm">missile_trajectory.xlsm</a></td>
      <td align="center" valign="top"><a href="missile_trajectory_key.xlsm">missile_trajectory_key.xlsm</a></td>
    </tr>
  </tbody>
</table>
<br />

<?php 
require "../footer.php";
?>
</body>
</html>
