<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Excel VBA Primer</title>
<link href="../../nljstyles.css" rel="stylesheet" type="text/css" />
<link href="../../prism/prism.css" rel="stylesheet" />
</head>

<body>
<script src="../../prism/prism.js"></script>
<h1>Adding Controls to a Spreadsheet</h1>
<p>Controls are buttons, combo boxes, option buttons, etc. that can be added to a worksheet. You can associate VB code with controls that is executed each time the user clicks on the control.</p>
<h2><a name="controlsgroup" id="controlsgroup"></a>The Controls Group</h2>
<p>To add a control to a spreadsheet, you should first turn on the <a href="../vb-gettingstarted/index.php#developertab">Developer tab</a>. You then create new controls and edit the code associated with controls using the Control group.</p>
<p><img src="../vb-gettingstarted/controlsgroup.png" width="174" height="87" /></p>
<p>The options in the group are as follows:</p>
<table border="0" width="100%" cellpadding="4" cellspacing="4">
  <tr>
    <td align="center"><img src="insert.gif" width="41" height="71" /></td>
    <td width="87%">This tool is used to insert new controls (see below).</td>
  </tr>
  <tr>
    <td align="center"><img src="designmode.gif" width="43" height="72" /></td>
    <td width="87%">This toggle button is used to switch in and out of design mode (see below)</td>
  </tr>
  <tr>
    <td align="center"><img src="properties.gif" width="83" height="21" /></td>
    <td width="87%">This button brings up the Properties window that is used to edit the properties associated with a selected control.</td>
  </tr>
  <tr>
    <td align="center"><img src="viewcode.gif" width="84" height="25" /></td>
    <td width="87%">This tool displays the <a href="../vb-gettingstarted/index.php#vbeditor">Visual Basic
    Editor</a>.&nbsp; This is where you write the Visual Basic code.</td>
  </tr>
  <tr>
    <td align="center"><img src="rundialog.gif" width="81" height="23" /></td>
    <td width="87%">This is used to run any custom user forms you may have associated with the project.</td>
  </tr>
</table>
<h2><a name="creatingcontrols" id="creatingcontrols"></a>Creating Controls</h2>
<p>To add a new control, click on the Insert button and then select one of the control types in the <strong>ActiveX Controls</strong> portion of the tool palette:</p>
<p><img src="toolpalette.gif" width="167" height="227" /></p>
<p>Then drag a box on your worksheet where you want to place the control. While the control is still selected, you can use the <strong>Properties</strong> window to edit the properties associated with the control (the control name, the caption, etc.).</p>
<p>To edit the code associated
  with a control, simply double click on the control. This switches you over the VB Editor and puts the cursor in the subprocedure associated with the <strong>Click</strong> event for the control. Any code you enter here is automatically executed when the control is clicked.</p>
<h2><a name="designmode" id="designmode"></a>Design Mode</h2>
<p>When you are writing VB code and adding controls to a spreadsheet, there are
  two basic modes: <b> Design mode</b> and <b> Run mode</b>.&nbsp; In design mode,
  when you click on a button or a control, you can edit the properties of the
  control in the <a href="../vb-gettingstarted/index.php#vbeditor">VB Editor</a>.&nbsp; If you double click
  on a control, you can edit the code associated with the control.&nbsp; If you
  are in Run mode, when you click on the control, the code associated with the
  control is executed.&nbsp; When developing your spreadsheet, you will be moving
  in and out of Design mode.</p>
<h2>Types of Controls</h2>
<p>In this section we will describe each of the controls in the ActiveX Controls menu.</p>
<h3>Buttons <img src="control_button.png" width="14" height="9" alt=""/></h3>
<p>Buttons are also called <strong>Command Buttons</strong>. They are typically used as a simple mechanism to execute some VBA code, such as a macro you my have recorded.</p>
<p><img src="buttons_sample.png" width="265" height="49" alt=""/></p>
<p>After creating the button, you double-click the button in design mode and then write the code that is executed when the button is clicked. If your button brings up a custom user form, you should also put a set of elipses after the button caption (...).</p>
<h3>Combo Boxes <img src="control_combobox.png" width="16" height="16" alt=""/></h3>
<p>Combo boxes are sometimes called &quot;pop-up menus&quot;. They allow you to select a choice from a set of options.</p>
<blockquote>
  <p><img src="combobox_sample.png" width="122" height="90" alt=""/></p>
</blockquote>
<p>The choices presented to the user come from a range of cells somewhere on your spreadsheet. You assign the set of choices to the combobox by setting the <strong>ListFillRange</strong> property of the combo box. The range is reference likes you would in a spreadsheet (J2:J5, Sheet2!G4:G6, etc.).</p>
<p>The &quot;combo&quot; part of a combo box comes from the fact that rather than selecting an item from the list, you can type something in the box. If you want to restrict the user to only select something from the list (typical usage), you need to set the <strong>MatchRequired</strong> property to True.</p>
<h3>Check Boxes <img src="control_checkbox.png" width="14" height="14" alt=""/></h3>
<p>Checkboxes are used when you have an option or set of options that are true/false or yes/no and the options are independent. For example,</p>
<blockquote>
  <p><img src="checkbox_sample.png" width="205" height="121" alt=""/></p>
</blockquote>
<p>The text displayed corresponds to the <strong>Caption</strong> property. The <strong>Value</strong> property is a boolean (true/false) indicating the status of the control. This control is typically utilized with some kind of <strong>If</strong> statement. You typically don't do anything when the user selects the control, but when the user selects a button or something, you check the status of the control. For example,</p>

<pre><code class="language-vb">If chkCoseSlaw Then
	price = price + 2.6
End If</code></pre>

<p>If your options are dependent (you can only select one from the group), you should use either the list box, combo box, or option control instead.</p>
<h3>List Boxes <img src="control_listbox.png" width="16" height="11" alt=""/></h3>
<p>A list box is used when you have a large number of options and you need the user to select one option from the list. Example:</p>
<blockquote>
  <p><img src="listbox_sample.png" width="178" height="190" alt=""/></p>
</blockquote>

<p>The contents of the list come from a range of cells somewhere on your spreadsheet. The range is indicated by the <strong>ListFillRange</strong> property. To determine what item is selected, you can check the <strong>Value</strong> property. For the example shown above, the Value property would = &quot;Brigham Young&quot;. In many cases, what you really want to know is what item is selected by order in the list. You can do this using the <strong>ListIndex</strong> property. The index numbering starts at 0, so for the case shown above, the ListIndex property = <strong>2</strong>.</p>
<p>Note that if the list box is not big enough to show the entire list, a scroll bar is automatically added to the list to allow the user to scroll through the entire list.</p>
<h3>Text Boxes <img src="control_textbox.png" width="16" height="13" alt=""/></h3>
<p>A text box is used to allow the user to enter a text string as some type of input. For example,</p>
<blockquote>
  <p><img src="textbox_sample.png" width="241" height="35" alt=""/></p>
</blockquote>
<p>The text in the control is stored in the <strong>Value</strong> property. Text boxes are rarely used in Excel because it is usually simpler and more efficient to have the user enter the text into one of the spreadsheet cells.</p>
<h3>Scroll Bars <img src="control_scroll.png" width="9" height="16" alt=""/></h3>
<p>A scroll bar allows the user to select a value from a predefined range of values. </p>
<blockquote>
  <p><img src="scrollbar_sample.png" width="207" height="219" alt=""/></p>
</blockquote>
<p>The rectangle in the middle is called the &quot;thumb&quot;. As the user moves the thumb, the value of the scroll bar changes and the value is stored in the <strong>Value</strong> property. You can link a scroll bar to a cell using the <strong>LinkedCell</strong> property. Then the current value of the scroll is echoed to a cell as shown above. You set the range of values on the scroll bar using the <strong>Min</strong> and <strong>Max</strong> properties. Scroll bars are fun, but are rarely used on spreadsheets.</p>
<h3>Spin Buttons <img src="control_spin.png" width="15" height="15" alt=""/></h3>
<p>A spin button allows the user to increment an integer value up or down by clicking on the button arrows. A spin button should always be linked to a cell using the <strong>LinkedCell</strong> property.</p>
<blockquote>
  <p><img src="spin_sample.png" width="223" height="37" alt=""/></p>
</blockquote>
<p>You can set a limit on the values by changing the <strong>Min</strong> and <strong>Max</strong> properties. You can also use the <strong>SmallChange</strong> property to increment by 2, 5, 10, etc (default = 1).</p>
<h3>Option Buttons <img src="control_option.png" width="10" height="10" alt=""/></h3>
<p>Option buttons are used when you want the user to make a single selection from a set of options. These are sometimes called &quot;radio&quot; controls, because the mimic the buttons on your car radio for selecting a pre-set station.</p>
<blockquote>
  <p><img src="option_sample.png" width="72" height="76" alt=""/></p>
</blockquote>
<p>Each control has a <strong>Value</strong> property which is a boolean representing the state of the control (true/false). The text displayed is from the <strong>Caption</strong> property.</p>
<h3>Labels <img src="control_label.png" width="9" height="9" alt=""/></h3>
<p>A label is simply a text string where you can change the value of the text string using VB code. These are sometimes useful in custom user forms, but are rarely used in spreadsheets because you can simply put a label in a cell and change the value of the cell.</p>
<h3>Images <img src="control_image.png" width="16" height="16" alt=""/></h3>
<p>An image control allows you to associate an image from a file with a control. The control is shown as a gray rectangle until you attach an image to the control using the <strong>Picture</strong> property. You select a file and the image in the file is then copied to the image control. You can use the <strong>Visible</strong> property to turn the image on or off and you can use the <strong>Top</strong> and <strong>Left</strong> properties to control the location of the image on your spreadsheet. You should only use this control on a spreadsheet if you intend to turn the image on or off or move it. Otherwise, just use the <strong>Insert|Pictures</strong> command in Excel.</p>
<p>See the <a href="../../homework/hw11/index.htm">Road Kill homework assignment</a> for an example of how to use image controls.</p>
<h3>Toggle Buttons <img src="control_togglebutton.png" width="12" height="16" alt=""/></h3>
<p>A toggle button is similar to a checkbox, but it uses a button that is in a normal or depressed state:</p>
<p><img src="togglebutton_sample.png" width="225" height="39" alt=""/></p>
<p>The button status is stored in the <strong>Value</strong> property as a boolean (true/false). The text on the button is from the <strong>Caption</strong> property.</p>
<h3>More Controls <img src="control_more.png" width="15" height="16" alt=""/></h3>
<p>Clicking on the more controls button brings up a list of advanced controls. The contents of the list will depend on what is installed on your computer. </p>
<p><img src="morecontrols.png" width="524" height="322" alt=""/></p>
<h2>Control Names</h2>
<p>When working with controls, you will need to reference your controls  in your code using the control names which are defined by the Name property when you create the controls. When dealing with large numbers of controls, it is very important to assign meaningful names that describe the function of the control. Furthermore, to distinguish betwen different types of controls, you should use a three character prefix on the control names to indicate the type. For example, if I had a set of checkboxes for BBQ sides as shown above, I would name them:</p>
<blockquote>
  <p>chkColeSlaw<br />
    chkPotatoSalad<br />
    chkBakeBeans<br />
    chkMacNCheese</p>
</blockquote>
<p>And for my option controls I would name them:</p>
<blockquote>
  <p>optRed<br />
    optGreen<br />
    optBlue</p>
</blockquote>
<p>Here is a complete list of suggested three character prefixes for each control type.</p>

  <table width="261" border="0">
    <tr>
      <th width="130" align="left" scope="col">Type</th>
      <th width="45" align="center" scope="col">Icon</th>
      <th width="72" align="center" scope="col">Prefix</th>
    </tr>
    <tr>
      <td align="left" >Command Button</td>
      <td align="center"><img src="control_button.png" width="14" height="9" alt=""/></td>
      <td align="center">cmd</td>
    </tr>
    <tr>
      <td align="left" >Combo Box</td>
      <td align="center"><img src="control_combobox.png" width="16" height="16" alt=""/></td>
      <td align="center">cbo</td>
    </tr>
    <tr>
      <td align="left" >Check Box</td>
      <td align="center"><img src="control_checkbox.png" width="14" height="14" alt=""/></td>
      <td align="center">chk</td>
    </tr>
    <tr>
      <td align="left" >List Box</td>
      <td align="center"><img src="control_listbox.png" width="16" height="11" alt=""/></td>
      <td align="center">lst</td>
    </tr>
    <tr>
      <td align="left" >Text Box</td>
      <td align="center"><img src="control_textbox.png" width="16" height="13" alt=""/></td>
      <td align="center">txt</td>
    </tr>
    <tr>
      <td align="left" >Scroll Bar</td>
      <td align="center"><img src="control_scroll.png" width="9" height="16" alt=""/></td>
      <td align="center">scr</td>
    </tr>
    <tr>
      <td align="left" >Spin Button</td>
      <td align="center"><img src="control_spin.png" width="15" height="15" alt=""/></td>
      <td align="center">spn</td>
    </tr>
    <tr>
      <td align="left" >Option Button</td>
      <td align="center"><img src="control_option.png" width="10" height="10" alt=""/></td>
      <td align="center">opt</td>
    </tr>
    <tr>
      <td align="left" >Label</td>
      <td align="center"><img src="control_label.png" width="9" height="9" alt=""/></td>
      <td align="center">lbl</td>
    </tr>
    <tr>
      <td align="left" >Image</td>
      <td align="center"><img src="control_image.png" width="16" height="16" alt=""/></td>
      <td align="center">img</td>
    </tr>
    <tr>
      <td align="left" >Toggle Button</td>
      <td align="center"><img src="control_togglebutton.png" width="12" height="16" alt=""/></td>
      <td align="center">tog</td>
    </tr>
  </table>

<?php 
require "../footer.php";
?>
</body>
</html>
