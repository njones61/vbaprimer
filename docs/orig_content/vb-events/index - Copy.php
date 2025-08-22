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
<h1>Trapping for Worksheet Events</h1>
<p>When writing VB code associated with a spreadsheet, it is common to add a 
  button to the spreadsheet that the user pushes to execute the VB code when the 
  desired changes have been made to the controls and the values have been entered 
  in the cells. For example, the following code performs some simple tests 
  to see if someone is ready to take the Professional Engineer test:</p>
<p><img src="../events/pequalfig1.jpg" width="704" height="392" /></p>
<p>In this case, we don't need any code for the click events for the option and 
  checkbox controls. We simply need to add the following code for the &quot;Do I 
  Qualify&quot; button:</p>
  
<pre><code class="language-vb">Private Sub cmdQualify_Click()

Dim qualify As Boolean
Dim nextrow As Integer

qualify = True
nextrow = 15

Range("B15:B17").ClearContents

If Not optFEYes Then
    qualify = False
    Cells(nextrow, 2) = "You did not pass the FE exam"
    nextrow = nextrow + 1
End If

If Range("years") < 4 Then
    qualify = False
    Cells(nextrow, 2) = "You do not have enough years of work experience."
    nextrow = nextrow + 1
End If

If chkUteGrad Then
    qualify = False
    Cells(nextrow, 2) = "You graduated from the wrong school!"
    nextrow = nextrow + 1
End If

If qualify Then
    Range("result") = "Congratulations! You qualify."
Else
    Range("result") = "Sorry! Keep trying."
End If

End Sub
</code></pre>
  
<p>This code works fine, but why require the user to click on the button? Why not set up the spreadsheet so that anytime the user clicks on a control or 
  changes the value of a cell, the VB code is automatically executed and the 
  results are updated? This can be easily accomplished using the &quot;Change&quot; 
  event for the worksheet.</p>
<h2>Workbooks and Worksheets</h2>
<p>Before discussing the Change event, we need to first define a couple of 
  terms. When the VB compiler for Excel is open, you will see a list of 
  objects in a tree on the left side of the window. At the bottom of the 
  tree you will see the following objects:</p>
<p><img border="0" src="tree.jpg" width="162" height="85" /></p>
<p>A &quot;Workbook&quot; object represents the entire spreadsheet, including all of the 
  sheets. If you double click on this object, it will bring up the source 
  code related to the workbook as a whole. The other objects (&quot;Sheet1&quot;, 
  &quot;Sheet2&quot;, &amp; &quot;Sheet3&quot;). Double clicking on these objects brings up the code 
  related to these objects.</p>
<p>Once you open the window related to a particular sheet, some important 
  information related to the sheet is displayed at the top of the sheet as 
  follows:</p>
<p><img src="sheetwindow.jpg" width="491" height="287" alt=""/></p>
<p>The combo box on the left (the one that is open) lists all of the objects 
  associated with the sheet. Note that each of the controls on the 
  spreadsheet are listed along with the worksheet itself. If you highlight 
  one of the objects, you can then select an event from the combo box on the 
  right:</p>
<p><img src="sheetwindow2.jpg" width="674" height="345" alt=""/></p>
<p>Selecting one of these events creates the subroutine for the selected event. For example, if I click on the &quot;Activate&quot; item, the following code appears:</p>

<pre><code class="language-vb">Private Sub Worksheet_Activate()

End Sub</code></pre>

<p>Any code inside this sub would be executed each time the associated sheet is made active.</p>

<h2>The Calculate and Change Events</h2>
<p>Note that the list of available events for the worksheet include the 
  &quot;Calculate&quot; event and the &quot;Change&quot; event. By selecting these items, we can 
  then fill in the code for these events. The resulting code will be 
  executed as follows:</p>
<h3>Calculate Event</h3>
<p>The Calculate event looks like this:</p>

<pre><code class="language-vb">Private Sub Worksheet_Calculate()

End Sub
</code></pre>


<p>and is called each time the formulas in the worksheet are recalculated. Note that you must have at least one formula in your spreadsheet in order for 
  this event to be called.</p>
<h3>Change Event</h3>
<p>The change event looks like this:</p>

<pre><code class="language-vb">Private Sub Worksheet_Change(ByVal Target As Range)

End Sub
</code></pre>

<p>and is called each time any of the cells in the spreadsheet are changed. Note that the subroutine takes one argument which is the range that has been 
  changed. If we want the spreadsheet to be updated any time the user enters 
  new data, this is the event we want to use. First of all, we remove the 
  button so that the spreadsheet looks as follows:</p>
<p><img src="../events/pequalfig3.jpg" width="714" height="409" /></p>
<p>Next, we will modify the code in the Change event to update the spreadsheet. However, this event is not called when a control is changed, it is only called 
  when a cell is changed. Therefore, we will first create a subroutine that 
  performs the calculations:</p>
  
<pre><code class="language-vb">Private Sub update_results()

Dim qualify As Boolean
Dim nextrow As Integer

qualify = True
nextrow = 15

Range("B15:B17").ClearContents

If Not optFEYes Then
    qualify = False
    Cells(nextrow, 2) = "You did not pass the FE exam"
    nextrow = nextrow + 1
End If

If Range("years") < 4 Then
    qualify = False
    Cells(nextrow, 2) = "You do not have enough years of work experience."
    nextrow = nextrow + 1
End If

If chkUteGrad Then
    qualify = False
    Cells(nextrow, 2) = "You graduated from the wrong school!"
    nextrow = nextrow + 1
End If

If qualify Then
    Range("result") = "Congratulations! You qualify."
Else
    Range("result") = "Sorry! Keep trying."
End If

End Sub
</code></pre>  
  

<p>Next, we will modify the Change event so that it calls this subroutine:</p>

<pre><code class="language-vb">Private Sub Worksheet_Change(ByVal Target As Range)

Application.EnableEvents = False

update_results

Application.EnableEvents = True

End Sub
</code></pre>

<p>Note that we have to temporarily turn off trapping for events prior to 
  updatings the results. This is because the Update_Results sub changes the 
  value of some of the cells. This generates a new Change event which brings 
  us right back to this sub. This results in an infinite loop. To be 
  safe, you should always turn off event trapping while you make any changes in 
  code.</p>
<p>Finally, to ensure that the click events for the controls cause the results 
  to be updated, we add a call to the click event subroutines for each of the 
  controls as follows:</p>
 
<pre><code class="language-vb">Private Sub optFEYes_Click()
update_results
End Sub

Private Sub optFENo_Click()
update_results
End Sub

Private Sub chkUteGrad_Click()
update_results
End Sub
</code></pre> 
 

<p>At this point, clicking on any of the controls, or updating the value in the 
  &quot;years&quot; cell triggers the VB code to update the results.</p>
<h2>Checking on the Target</h2>
<p>Note that Worksheet Change event sends an argument called Target that 
  represents the range of cells changed. This could be a single cell or a 
  range of cells. In some cases, it is useful to check on the range of cells 
  that have been modified. To do this, you can check on the Target object 
  passed as a parameter to the Change event sub. Target contains the cell or 
  range of cells changed. Ideally, you could use a simple statement such as:</p>

<pre><code class="language-vb">If Target = Range("B4") Then
</code></pre>
  
<p>or</p>

<pre><code class="language-vb">If Target <> Range("B13") Then
</code></pre>

<p>To check the value of Target. However, this will not work because both 
  Target and Range() are objects. As an object, when you say</p>
  
<pre><code class="language-vb">If Target = Range("B4") Then
</code></pre>

<p>what you are really saying is</p>

<blockquote>
  <p>&quot;if the value of target is equal to the value of cell B4, then&quot;</p>

</blockquote>
<p>This statement would return true if Target corresponded to <b>ANY</b> cell 
  that happened to have the same value as cell B4. A simple way to solve 
  this problem is to check on the <b>Address</b> property as follows:</p>
  
<pre><code class="language-vb">If Target.Address = Range("B4").Address Then
</code></pre>  
  

<p>In many cases, however, what you really want to know is whether or not Target 
  is a portion of an entire range of cells. An efficient test for this type 
  of case is to use the <b>Intersect</b> method associated with the <b>Application</b> object. This method returns the intersection between two ranges. The 
  idea is to intersect the target range and the range corresponding to the input 
  cells and see if the result is non-empty. This can be accomplished as 
  follows:</p>

<pre><code class="language-vb">Private Sub Worksheet_Change(ByVal Target As Range)

If Not Application.Intersect(Target, Range("B6")) Is Nothing Then
    update_events
End If

End Sub
</code></pre>  
  

<p>This approach will work with input ranges spanning multiple cells. For 
  example:</p>
  
<pre><code class="language-vb">Private Sub Worksheet_Change(ByVal Target As Range)

If Not Application.Intersect(Target, Range("B4:F34")) Is Nothing Then
    update_events
End If

End Sub
</code></pre>  
<br />
<?php 
require "../footer.php";
?>
</body>
</html>
