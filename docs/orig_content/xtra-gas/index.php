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
<h1>Google Sheets and Google Apps Script</h1>
<p>Now that you are familiar with Excel and VBA, you may wish to start exploring Google Sheets and Google Apps Script. When you create a Google account, you have free access to cloud storage on Google Drive, where you can use a free suite of productivity tools to create documents. One of these free tools is Google Sheets, a spreadsheet application. Google Sheets is very similar to Excel. Formulas are created in the same fashion and most of the same functions are supported. </p>
<p>Google Sheets supports a powerful scripting environment called Google Apps Script (GAS). GAS is a variant of JavaScript, the most common scripting language for web pages. GAS is an object-oriented programming language. A complete reference for the classes used in the Spreadsheet service of GAS can be found here:</p>
<blockquote>
  <p><a href="https://developers.google.com/apps-script/reference/spreadsheet/" target="_blank">https://developers.google.com/apps-script/reference/spreadsheet/</a></p>
</blockquote>
<p>Example code is included. You can also google countless sites with sample GAS code and problem solutions.</p>
<h3>Google Sheets vs Excel</h3>
<p>Here are some differences/comparisons between Google Sheets/GAS and Excel/VBA:</p>
<blockquote>
  <p><strong>Free</strong> - Excel is a commercial product. Google Sheets is free.</p>
  <p><strong>Multi-Platform</strong> - While the Mac version of Excel supports VBA, it is limited. Google Sheets works on all platforms exactly the same way. All you need is a browser.</p>
  <p><strong>Script Security</strong> - With Excel, you have to mess with the Macro security settings to get your VBA code to work and you have to save the file as *.xlsm. With Google Sheets you simply give it permission to run once by making a few clicks and that is it. </p>
  <p><strong>Sharing</strong> - With Google Sheets you can easily share the document with someone else and each of you can edit it, even at the same time. Fantastic for collaborative projects.</p>
  <p><strong>Ease of Scripting</strong> - Here I would give the edge to Excel/VBA. GAS has a steeper learning curve. But don't let that stop you. Once you get the hang of things you will be turning out code like a champ.</p>
  <p><strong>Controls</strong> - Google Sheets does NOT support ActiveX controls. You can execute your GAS in one of four ways: a) Using a menu command, b) using a drawing object (you can make it look like a button,  c) as a custom function in a formula, d) using the run button in the Script Editor window.</p>
  <p><strong>Recording Macros</strong> - This is a tie. Both platforms allow you to record actions and turn them into code. </p>
</blockquote>
<h3><strong>Learning Javascript</strong></h3>
<p>Before diving into Google Sheets and GAS, I recommend you spend some time learning about Javascript. No need to buy a textbook, there are tons of free resources on the web. I recommend the following site:</p>
<blockquote>
  <p><a href="http://www.w3schools.com/" target="_blank">http://www.w3schools.com/</a></p>
</blockquote>
<p>Click on the JavaScript link on the left.</p>
<h3>Importing Excel Files</h3>
<p>If you have an Excel file you want to try in Google Sheets, just upload it to your Google Drive and then right-click on it and select Open With|Sheets. It will create a copy of the file in Google Sheets format and open it. Your VBA code will not be preserved, but almost everything else will be. </p>
<h3>Opening a New Sheet</h3>
<p>You can also create a new blank sheet. In Google Drive, click on the New button and then select Google Sheets.</p>
<p><img src="new_sheet.png" width="346" height="436" alt=""/></p>
<h3>Opening the Editor</h3>
<p>Once you open your sheet, you can access the editor by selecting the <strong>Tools|Script Editor... </strong>command. This takes you to the editor with a new project, a new code file (Code.gs), and an empty function:</p>
<p><img src="new_project.png" width="525" height="312" alt=""/></p>
<h3>Writing Your First Function</h3>
<p>Lets change the function code so that it multiplies the input by 2:</p>
<blockquote>
  <p class="code">function double_it(x) {<br />
    &nbsp;&nbsp;return x*2;<br />
    }</p>
</blockquote>
<p>Save the changes and then go to the sheet and try the formula:</p>
<p><img src="double_it_1.png" width="371" height="406" alt=""/></p>
<p>It's that easy!</p>
<p><img src="double_it_2.png" width="324" height="393" alt=""/></p>
<h3>Hello World</h3>
<p>Let's write a function that prints &quot;hello world&quot;. In Excel VBA, we would write a custom sub. In GAS, we write a function with no parameters and it behaves like a sub. Add the following:</p>
<blockquote>
  <p class="code">function hello_world() {<br />
&nbsp;&nbsp;var ss = SpreadsheetApp.getActiveSpreadsheet();<br />
&nbsp;&nbsp;var sheet = ss.getSheetByName(&quot;Sheet1&quot;); <br />
&nbsp;&nbsp;sheet.getRange(&quot;C3:D11&quot;).setValue(&quot;hello world&quot;);<br />
    }</p>
</blockquote>
<p>To run the code, change the function selector in the toolbar to the editor to &quot;hello_world&quot;. This sets the active function.</p>
<p><img src="hello_world_1.png" width="583" height="304" alt=""/></p>
<p> Then click on the Play button. The first time you do this you will need to give permission for the script to run. When it finishes, you should see this:</p>
<p><img src="hello_world_2.png" width="513" height="425" alt=""/></p>
<h3>Recording a Macro</h3>
<p>Next, let's record a macro. This process is almost identical to Excel. We will record a simple macro that formats the range of &quot;hello world&quot; cells. First, select the <strong>Macros|Record macro</strong> command in the <strong>Tools</strong> menu.</p>
<p><img src="macro_1.png" width="759" height="416" alt=""/></p>
<p>This puts you into <em>record</em> mode. You should see this window at the bottom of your screen:</p>
<p><img src="macro_2.png" width="497" height="209" alt=""/></p>
<p>Next, apply some formatting as follows:</p>
<p>1) Select all of the cells that contain &quot;hello world&quot;<br />
  2) Change the font to italic<br />
3) Change the cell alignment to center the text<br />
4) Apply borders to the selected cells<br />
5) Fill the selected cells with a color</p>
<p>You don't have to follow those steps exactly. Feel free to apply whatever formatting you want. When you are done it will look something like this:</p>
<p><img src="macro_4.png" width="485" height="418" alt=""/></p>
<p>When you are done, select the <strong>Save</strong> command in the macro window at the bottom of your browser. That will bring up the save dialog. Enter &quot;<strong>my_macro</strong>&quot; and select <strong>Save</strong>.</p>
<p><img src="macro_3.png" width="418" height="316" alt=""/></p>
<p>Now we are ready to try our macro out. Before doing so, select the <strong>undo</strong> button (repeatedly if necessary) to remove the formatting.</p>
<p><img src="macro_5.png" width="203" height="191" alt=""/></p>
<p>To run the macro, select the <strong>Macros|my_macro</strong> command.</p>
<p><img src="macro_6.png" width="725" height="414" alt=""/></p>
<p>Boom! Your table should be reformatted. </p>
<p>To view the code, go back to the Code Editor and click on the macros.gs item on the left. Your macro code will appear. As is the case with Excel, you can now modify this code, copy-paste it to a different module, etc.</p>
<p><img src="macro_7.png" width="1127" height="303" alt=""/></p>
<h3>Finished!</h3>
<p>OK, now you have enough to get started. Have fun!</p>
<p>Click <a href="https://docs.google.com/spreadsheets/d/1YXzBVI6e-W3IXEoUfOnZfIAUzdcA6ARqIq1DFfWzg-s/copy">here</a> to get a copy of the sheet used in this page.</p>
</body>
</html>
