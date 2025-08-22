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
<h1>Debugging - Finding and Fixing Errors</h1>
<p>One of the more frustrating parts of programming can be finding and fixing errors. This process is called &quot;debugging&quot; and errors in computer code are called &quot;bugs&quot;. Debugging your code is much easier if you utilize the built-in debugging tools provided with the Visual Basic Editor.</p>
<h2>Types of Bugs</h2>
<p>Before discussing the debugging tools, it is helpful to review the three basic types of bugs: syntax errors, run-time errors, and logic errors.</p>
<h3>Syntax Errors</h3>
<p>All programming languages have a set of rules that must be followed when writing the code. These rules are called the code syntax. When you write code, a utility called the compiler parses through your code and converts it to low-level machine language. If you do not use proper syntax, the compiler cannot understand what it is you are trying to do. For example, a common syntax error would be be forgetting to include the &quot;Then&quot; part of an If statement. If you hit the return key before typing &quot;Then&quot;, you get a message like this:</p>
<blockquote>
  <p><img src="syntax_message.png" width="364" height="229" alt=""/></p>
</blockquote>
<p>Another common error would be mispelling a keyword (&quot;Iff&quot;). Most syntax errors are caught immediately by the VB Editor when you hit the return key. You can also check for syntax errors by selecting the <strong>Compile VBA Project</strong> command in the <strong>Debug</strong> menu.</p>
<p>Syntax errors are the easiest to find and fix.</p>
<h3>Run-Time Errors</h3>
<p>Run-Time errors occur when your code has correct syntax, but the logic of the code is such that it generates an error when you execute the code. For example, the code:</p>

<pre><code class="language-vb">x = 0
y = 1 / x 
</code></pre>


<p>Has legal syntax but generates a &quot;Divide By Zero&quot; error when you run the code. When this happens you get an error message and the Visual Basic Editor pops to the front and the line with the offending code is highlighted. Sometimes it is obvious what is causing the run-time error, but in some cases you will need to use the debugging tools to figure it out.</p>
<h3>Logic Errors</h3>
<p>A logic error is when your code has correct syntax and does not produce a run-time error but your code produces incorrect results (generates the wrong numbers, for example). Logic errors are the most difficult kind of bug to find and fix. You can spend hours trying to fix your logic errors, but the debugging tools described below can greatly simplify the process.</p>
<h2>Debugging Tools</h2>
<p>The debugging tools are located in the Debug menu in the VB Editor. These tools can be used to  trace the excution of your code in a line-by-line fashion and examine the value of variables to determine what is causing your run-time or logic errors. </p>
<h3>Breakpoints</h3>
<p>The first thing you should do when debugging is create one or more breakpoints. A breakpoint is a location in the code where you want execution to halt so that you can examine the state of your variables and objects and/or begin to trace execution line-by-line. To set a breakpoint, you simply click in the gray column on the left side of the editor and a red dot will appear:</p>
<blockquote>
  <p><img src="breakpoint.png" width="371" height="149" alt=""/></p>
</blockquote>
<p>When you run the code, the execution will then stop at the breakpoint and the line will be highlighted in yellow.</p>
<blockquote>
  <p><img src="breakpoint2.png" width="380" height="149" alt=""/></p>
</blockquote>
<p>The yellow highlight means that the yellow line has NOT been executed yet, but it is the next line to execute.</p>
<p>You can also set a breakpoint by putting the cursor on a line and selecting the <strong>Debug|Toggle Breakpoint</strong> command or by pressing <strong>F9</strong>. You can clear a breakpoint by clicking on it. You can get rid of all of your breakpoints by selecting the <strong>Debug|Clear All Breakpoints</strong> command.</p>
<h3>Step Options</h3>
<p>Once you have halted execution using a breakpoint, you can then trace the execution of your code one line at a time using the Step options. In most cases, the easiest thing to do is press the <strong>F8</strong> key to execute the <strong>Step Into</strong> command. This executes the current line and then highlights the next line in yellow. However, if the current line contains a custom sub or function (i.e., one that you wrote as opposed to a built-in VB sub or function), then you need to be a little more careful. There are actually three different Step commands and each behaves a little differently based on how you want to trace your code.</p>
<table width="747" border="0">
  <tr>
    <td width="133"><strong>Command</strong></td>
    <td width="120" align="left"><strong>Shortcut</strong></td>
    <td width="480"><strong>Result</strong></td>
  </tr>
  <tr>
    <td valign="top">Step Into</td>
    <td align="left" valign="top">F8</td>
    <td valign="top">Executes the current line of code and if the code contains a call to a custom sub or function, the code execution jumps to the first line in the sub or function and then pauses. This gives you the chance to trace the execution of the sub or function.</td>
  </tr>
  <tr>
    <td valign="top">Step Over</td>
    <td align="left" valign="top">Shift+F8</td>
    <td valign="top">Executes the current line of code and if the code contains a call to a custom sub or function, the sub or function is executed but that execution is not traced. After executing the line (including referenced subs and/or functions) the next line of code in the current module is highlighted and the execution pauses. You should use this option when you are confident that your subs or functions do not contain errors and you don't need to trace their execution.</td>
  </tr>
  <tr>
    <td valign="top">Step Out</td>
    <td align="left" valign="top">Ctrl+Shift+F8</td>
    <td valign="top">This command should be used when you are in the middle of tracing the execution of your code inside a custom sub or function and you are satisfied that the sub or function does not contain the error you are looking for. This completes the execution of the sub or function and returns you to the next line following the line where the sub or function was called.</td>
  </tr>
  <tr>
    <td valign="top">Run to Cursor</td>
    <td align="left" valign="top">Ctrl+F8</td>
    <td valign="top">This command executes all code between the current yellow-highlisted line and the line containing the cursor. The line containing the cursor is then highlighted and the execution pauses.</td>
  </tr>
</table>
<p>It is important to note that none of these four options &quot;skips&quot; code in the sense that portions of your code are not executed. All of the code is executed, but in some cases you may not see the step-by-step execution. If for some reason you want to actually skip one or more lines of code, you should set the cursor in the first line following the code you want to skip and then select the <strong>Debug|Set Next Statement</strong> command.</p>
<h3>Run Commands</h3>
<p>After stopping at a breakpoint and examining some code, you may wish to quickly finish all of your remaining lines of code, or execute all lines of code between the current line and the next breakpoint. You can do this using the Run command. There is a set of Run-Pause-Stop commands in the VB Editor menu.</p>
<table width="537" border="0">
  <tr>
    <td width="87"><strong>Command</strong></td>
    <td width="78" align="center"><strong>Symbol</strong></td>
    <td width="358"><strong>Action</strong></td>
  </tr>
  <tr>
    <td valign="top">Run</td>
    <td align="center" valign="top"><img src="command_run.png" width="9" height="15" alt=""/></td>
    <td valign="top">Execute the code from the current position to the next breakpoint or until the code is completed.</td>
  </tr>
  <tr>
    <td valign="top">Pause</td>
    <td align="center" valign="top"><img src="command_pause.png" width="11" height="13" alt=""/></td>
    <td valign="top">Pause the execution. This button can be used when your code appears to be stuck in an infinite loop. It causes the code execution to pause and the next line of code to be executed is highlighted.</td>
  </tr>
  <tr>
    <td valign="top">Stop</td>
    <td align="center" valign="top"><img src="command_stop.png" width="13" height="13" alt=""/></td>
    <td valign="top">This command can be used either when the code is running or when it is paused. It causes the execution to stop and everything is reset to the non-running state.</td>
  </tr>
</table>
<p>The Run command can also be used as a shortcut to run a specific sub by putting the cursor in the sub and then selecting the Run button in the VB Editor. For example, this could be used to execute the code associated with a button without actually having to click the button.</p>
<h3>Examining Variables and Expressions</h3>
<p>When the code execution is paused, it is extremely helpful to examine the status of your variables, expressions, and objects. To examine the value of a variable, simply hover the cursor over the variable and the current value is displayed:</p>
<blockquote>
  <p><img src="cursor_hover_variable.png" width="368" height="66" alt=""/></p>
</blockquote>
<p>You can also select part or all of an expression to see the value of the highlighted part:</p>
<blockquote>
  <p><img src="cursor_hover_expression.png" width="322" height="66" alt=""/></p>
</blockquote>
<p>If you find yourself checking the value of a particular variable or expression repeatedly, you can select the variable or expression and then select the <strong>Debug|Add Watch...</strong> command. This opens the Watch window at the bottom of the VB Editor and displays the value of the variable or expression in the window.</p>
<blockquote>
  <p><img src="watches.png" width="587" height="140" alt=""/></p>
</blockquote>
<p>These values are constantly updated as the code is executed.</p>
<p>If you add an array or object to the watch window, you can click the plus button to expand the object and view the values of object/array members. In this example, x and y are two arrays of doubles:</p>
<blockquote>
  <p><img src="watch_arrays.png" width="589" height="311" alt=""/></p>
</blockquote>

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
      <td> <strong>Total Head Debug -</strong> Find a run-time error, logic error and syntax error for the code solving for the total head in the bernoulli equation.</td>
      <td align="center" valign="top">Easy</td>
      <td align="center" valign="top"><a href="total_head_debug.xlsm">total_head_debug.xlsm</a></td>
      <td align="center" valign="top"><a href="total_head_debug_key.xlsm">total_head_debug_key.xlsm</a></td>
    </tr>
    <tr>
      <td><strong>Citation Machine - </strong>Debug the code so that a message box with the correct citation displays when the user clicks "Create Citation."</td>
      <td align="center" valign="top">Medium</td>
      <td align="center" valign="top"><a href="citation_machine.xlsm">citation_machine.xlsm</a></td>
      <td align="center" valign="top"><a href="citation_machine_key.xlsm">citation_machine_key.xlsm</a></td>
    </tr>
    <tr>
      <td><strong>Debugging -</strong> Debug the code that calculates three equations.</td>
      <td align="center" valign="top">Hard</td>
      <td align="center" valign="top"><a href="debugging.xlsm">debugging.xlsm</a></td>
      <td align="center" valign="top"><a href="debugging_key.xlsm">debugging_key.xlsm</a></td>
    </tr>
  </tbody>
</table>
<br />

<?php 
require "../footer.php";
?>
</body>
</html>
