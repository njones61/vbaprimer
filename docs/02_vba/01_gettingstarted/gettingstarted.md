# Getting Started - Using the Developer Interface and the VB Editor

Visual Basic (VB) is a great programming language for beginning programmers. It has a simple structure and it provides a number of safeguards that prevent common programming errors. Another great feature of VB is that it can be used as a powerful scripting language for writing macros and extensions to Microsoft Office applications including Excel and Access. It can also be used to write scripts for use in AutoCAD. There is a special version of VB used in these applications called Visual Basic for Applications (VBA).

Writing VBA code for Excel is easy and fun!! Once you learn a few basics, you will be creating highly professional spreadsheets. VBA allows you to design a spreadsheet that will do things that are impossible with the basic spreadsheet options. It also allows you to make your spreadsheets more user-friendly.

The Developer Tab
The first step in adding Visual Basic to your spreadsheet is to turn on the Developer tab. This is not a default part of the ribbon, so you may need to turn it on as follows:

Select the File|Options men command.
Click on the Customize Ribbon button on the left.
Turn on the Developer option shown in the Main Tabs section on the right.


Click OK to exit.
You should now see the Developer tab. This is where we interact with our VB code.



The Code Group
The Code group is used to record macros and to open the VB editor. The Visual Basic button opens the Visual Basic Editor window and the other tools are used to record and control macros.



The Visual Basic Editor
The VB Editor is where you edit the Visual Basic code.  It is very similar to the regular Visual Basic compiler.  The code is shown in a set of windows on the right.  The Project window on the left lists the components of the project.  The VBAProject folder lists each of the sheets in your spreadsheet and the workbook.  The Modules folder lists the code associated with Macros. The Forms folder lists the custom user forms associated with the project. To edit the code associated with a sheet, module, or user form, you simply double-click on the object in the Project Explorer Window.



The Controls Group
The Controls group is used to add controls to a worksheet and to create/edit the VB code associated with the controls.



The View Code button brings up the Visual Basic Editor window shown above.

Security Settings
Since VBA is such a flexible and powerful scripting environment, it also happens to be a popular method for writing viruses. For example, it is possible to write scripts that are automatically executed whenever a spreadsheet is opened. The script could theoretically attempt to do some damage to your computer (delete files, etc.) once it executes. To minimize the chance that a malicious script could cause damage, Microsoft turns on some default layers of security over VBA scripts. Before we can start writing VBA code, we need to adjust those settings.

Go to the Developer tab.
Click on Macro Security
You will then be presented with the following options:



Select the settings shown in the figure above and click OK.

You should only need to do this once. These settings are associated with your installation of Excel and will be applied each time you open a spreadsheet from here on out.