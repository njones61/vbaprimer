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
<h1>Looping Over Files</h1>
<p>One of the greatest benefits of using VBA with Excel is that you can automate tasks that can ordinarily be time-consuming. One form of automation that can be especially useful is to automatically open a set of files in a folder and open each of the files, make a change to the file content, and save the file. This can be accomplished quite easily with VBA in Excel, especially if the files correspond to Excel spreadsheets.</p>
<p>In this page, we will work through an example of modifying a set of spreadsheet files via VBA. Using this example as a guide, you can modify the code to fit your circumstances. The files associated with this exercise are in a zip archive and can be downloaded here:</p>
<blockquote>
  <p><a href="loopingfiles.xlsm">loopingfiles.xlsm</a><br />
  <a href="samplefiles.zip">samplefiles.zip</a> </p>
</blockquote>
<p> If you wish to follow along, please download and unzip the files in the zip archive to a folder named &quot;samplefiles&quot;. Then open the spreadsheet file (&quot;loopingfiles.xlsm&quot;) in Excel.</p>
<h2>Renaming the Files</h2>
<p>When you unzip the samplefiles.zip archive, you should see the following set of files:</p>
<blockquote>
  <p><img src="file_list.png" width="498" height="214" alt=""/></p>
</blockquote>
<p>In this case, each of these files is empty, but in other cases they may contain data. Our objective is to open each of these files and copy a table to the main sheet and then save the changes. \</p>
<p>Whenever you loop over files, you must have some systematic way of determining the names of the files in the folder. There used to be a FileSearch object that would list all of the files in a directory, but it was deprecated (discontinued) by Miscrosoft because it was being used to write viruses. So the simplest thing to do now is to name the files  so that we can formulate the file name in code as we iterate through a For loop from 1 to the number of files. Fortunately, it is rather easy to rename the files as shown. If you have an existing set of files to rename, you can do it as follows:</p>
<ol>
  <li>Select all of the files</li>
  <li>Right-click on the files and select the Rename command.</li>
  <li>Enter a common name (&quot;worksheet&quot; in the example shown above) and hit the return key.</li>
</ol>
<p>At this point all files are renamed as follows:</p>
<blockquote>
  <p><img src="file_list2.png" width="507" height="218" alt=""/></p>
</blockquote>
<p>This format is easier to recreate using code (see below).</p>
<h2>Input Options</h2>
<p>When you open the spreadsheet file you will see the main page:</p>
<blockquote>
  <p><img src="screenshot.png" width="711" height="521" alt=""/></p>
</blockquote>
<p>The inputs to the code are in three cells: <strong>B11</strong>, <strong>B13</strong>, and <strong>B15</strong>. B11 contains the path to the folder containing the files you wish to modify. B13 contains the prefix used when naming the files. Compare to file list shown above. B15 contains the number of files. Please note that cells B11, B13, and B15 have been named <strong>folderlocation</strong>, <strong>prefix</strong>, and <strong>nfiles</strong>, respectively. You may need to change these values before proceeding.</p>
<p>Before looking at the code, click on the Sheet2 tab and note the contents:</p>
<blockquote>
  <p><img src="screenshot2.png" width="441" height="406" alt=""/></p>
</blockquote>
<p>For our example problem, we will be copying this table from the loopingfiles.xlsm workbook to the first sheet in each of the files in the samplefiles folder and saving the changes.</p>
<h2>Code</h2>
<p>Next we will look at the source code associated with the <strong>Fix Files</strong> button.</p>

<pre><code class="language-vb">Private Sub cmdFixFiles_Click()
Dim myrow As Integer
Dim i As Integer
Dim nfiles As Integer
Dim filepath As String

'Set the default working directory
ChDir Range("folderlocation")

'Loop through each of the files in the folder
nfiles = Range("nfiles")
For i = 1 To nfiles

    'Copy the header and table on Sheet2 to the clipboard
    Sheets("Sheet2").Range("A1:E18").Copy

    'Formulate a text string identifying the full path to file i
    filepath = Range("startpath") & "\" & Range("prefix") & " (" & i & ").xlsx"
    
    'Open the file
    Workbooks.Open Filename:=filepath

    'Select the upper left cell and past the clipboard contents
    ActiveSheet.Range("A1").Select
    ActiveSheet.Paste
    
    'Fit the column widths
    ActiveSheet.Columns("B:B").EntireColumn.AutoFit
    ActiveSheet.Columns("C:C").EntireColumn.AutoFit
    ActiveSheet.Columns("D:D").EntireColumn.AutoFit
    ActiveSheet.Columns("E:E").EntireColumn.AutoFit
    
    'Select one of the cells so that the entire table is no longer selected (optional)
    ActiveSheet.Range("B4").Select
    
    'Save and close the file
    ActiveWorkbook.Save
    ActiveWorkbook.Close

    'Go to the next file
Next i

'Exit cut/copy mode (optional)
Application.CutCopyMode = False

End Sub
</code></pre>


<p>Each of the steps in the code is documented with a comment. Note how the folder location, prefix, and file number are used to generate a complete path to the file as shown on this line:</p>

<pre><code class="language-vb">'Formulate a text string identifying the full path to file i
filepath = Range("startpath") & "\" & Range("prefix") & " (" & i & ").xlsx"
</code></pre>

<p>Also, note that once you open another workbook, you have to be very careful how you reference cells and ranges. For example, if you reference cell <strong>A4</strong>, to which workbook does that apply? To ensure that there is no confusion, you should add the <strong>ActiveSheet</strong> or <strong>ActiveWorkbook</strong> prefix to all references to the external workbook after you open it. If ou need to refer to the current workbook (the one containing the code) while the other workbook is open, use the prefix <strong>ThisWorkbook</strong> before all sheet or range references.</p>
<p>If you are not sure how to structure the intercation between your two workbooks, you can always record a macro and perform the steps you wish to perform and then examine the macro code and adapt it to the sample shown above.</p>
<p>You may wish to try running the code above. You will see the screen flash once for each sample file as the code runs. After running the code, open each of the sample files to verify that the table was properly copied.</p>
<?php 
require "../footer.php";
?>
</body>
</html>
