# VBA Coding Standards

In this chapter we discuss practices that will make your code easier to read and understand and less prone to error. These are not strictly required - your code may run just fine when you do not follow these guidelines. However, it is strongly suggested that you get in the habit of using good coding practices as it will make your coding much simpler and enjoyable in the long run. This is list is not comprehensive, but is a good start.

## Option Explicit

As described in the [Variables](../06_variables/variables.md) chapter, you don't necessarily have to declare your variables before you use them. For example, if I want to start using a variable called **myvar**, I can simply write a line of code like this:

```vb
myvar = Range("A4") * 2
```

and my code will work just fine. When that line of code is executed, the VBA compiler looks at the myvar item and if it doesn't recognize it as a known variable or object, it assumes that you are declaring a new variable. That works fine until you write a line of code that accesses an existing variable but you misspell the variable name, thus accidentally creating a new variable with a default value of zero (or empty string, etc.). This leads to run-time errors that can be difficult to find and fix. This can all be avoided by getting into the habit of typing:

```vb
Option Explicit
```

at the top of your code every time you start coding a new module or sheet.

## Indenting

Indenting is not required in VBA, but it makes your code much easier to read and understand. For example, suppose you have a loop within a loop and the inner loop contains an IF statement, like this:

```vb
For myrow = 1 To 20
  For mycol = 5 To 15
    x = Cells(myrow, mycol)
    If (x < 0) Then
      MsgBox "Negative number found"
      numneg = numneg + 1
    ElseIf (x = 0) Then
      MsgBox "Zero value found"
      numzero = numzero + 1
    End If
  Next mycol
Next myrow
```

Notice how the indenting illustrates the flow of logic? Everything between the For myrow = and the Next myrow statements is indented to illustrate that the indented part is executed at each iteration of the loop. Likewise, the indentation for the inner loop and each section of the IF statement clearly identify the flow of logic. Now let's look at the same code without indentation:

```vb
For myrow = 1 To 20
For mycol = 5 To 15
x = Cells(myrow, mycol)
If (x < 0) Then
MsgBox "Negative number found"
numneg = numneg + 1
ElseIf (x = 0) Then
MsgBox "Zero value found"
numzero = numzero + 1
End If
Next mycol
Next myrow
```

This code will generate precisely the same results as the first code when executed. However, it is incredibly difficult to follow the logic. Thus, it is much easier to make mistakes when you do not indent. It is important to get into the habit of indenting your code. Just follow the example used in this Primer. We always indent lines of code that are in a logical block.

## Comments

When writing complex code, it is helpful to use comments to illustrate the logic in your code. You don't need to comment everything, but adding some comments as appropriate to explain the logic can make your code much easier to follow. This helps others who review your code and it may help you when you come back to your code at a later point in time. As much as you think you will remember everything, it is easy to forget what you were doing. For example, here is some well-commented code from a term project submitted for this course:

```vb
Private Sub cmd_generatecanoe_Click()

'get user input values for canoe parameters and convert entered values to feet and inches
Dim tl As Double
Dim tw As Double
Dim th As Double
tl = lst_tlfeet + lst_tlinches / 12
tw = lst_twfeet + lst_twinches / 12
th = lst_thfeet + lst_thinches / 12

'print values on reference data sheet
Sheets("Reference Data").Range("F2") = tl
Sheets("Reference Data").Range("F3") = tw
Sheets("Reference Data").Range("F4") = th

'Create array to generate X and Y points
'There are 31 points, which yields 11 nodes after drawing the curve
Dim canoeArray1(1 To 31, 1 To 2) As Single

'set up values for coordinate transformation
a = 20
c = -20
d = 150
b = 100

'set X-values evenly spaced across the canoe
For i = 1 To 31
    canoeArray1(i, 1) = (i - 1) * tl / 30
Next i

'Set Y values based on a sqrt function of the X value
'This first loop accounts for the Y-values increasing up to the center point
For i = 1 To 16
    'set y value
    canoeArray1(i, 2) = Sqr(tw ^ 2 * (canoeArray1(i, 1) / (tl / 2)))
    'Print X and Y values in feet/inches to spreadsheet for reference
    Sheets("All Canoe Points").Range("B" & i + 1) = canoeArray1(i, 1)
    Sheets("All Canoe Points").Range("C" & i + 1) = canoeArray1(i, 2)
    'Coordinate transformation
    canoeArray1(i, 1) = canoeArray1(i, 1) * a + b
    canoeArray1(i, 2) = canoeArray1(i, 2) * c + d
    'Print X and Y values in excel coordinates to spreadsheet
    Sheets("All Canoe Points").Range("G" & i + 1) = canoeArray1(i, 1)
    Sheets("All Canoe Points").Range("H" & i + 1) = canoeArray1(i, 2)
Next i

'This second loop accounts for the Y values decreasing from the center point on
For i = 17 To 31
    'set y value
    canoeArray1(i, 2) = Sqr(tw ^ 2 * (tl - canoeArray1(i, 1)) / (tl / 2))
    'Print X and Y values in feet/inches to spreadsheet for reference
    Sheets("All Canoe Points").Range("B" & i + 1) = canoeArray1(i, 1)
    Sheets("All Canoe Points").Range("C" & i + 1) = canoeArray1(i, 2)
    'coordinate transformation
    canoeArray1(i, 1) = canoeArray1(i, 1) * a + b
    canoeArray1(i, 2) = canoeArray1(i, 2) * c + d
    'Print X and Y values in excel coordinates to spreadsheet for reference
    Sheets("All Canoe Points").Range("G" & i + 1) = canoeArray1(i, 1)
    Sheets("All Canoe Points").Range("H" & i + 1) = canoeArray1(i, 2)
Next i

'draw bird's eye view of half of canoe
Set Sh = ActiveSheet.Shapes.AddCurve(canoeArray1)
'name the bird's eye view canoe1
Sh.Name = "canoe1"
```

Notice how the comments make the code easy to follow. This is highly commented code, and maybe a little on the overkill side, but it is helpful. Knowing how many comments to use is something you will get better at with experience but zero comments will definitely make your code difficult to understand!

## Control Names

As described in the [Controls](../03_controls/controls.md) chapter, using descriptive names for your controls is very important because it allows you to keep track of both the control's type and its function. You should always use descriptive control names and not use the default names.

## Variable Names

By the same logic, you should be careful when selecting your variable names. You should generally select variable names that accurately describe the intended use and contents of your variable. For example, suppose I am declaring an integer variable that will contain the number of students in a class. I could declare it as:

```vb
Dim ns As Integer
```

where ns stands for "number of students". A better approach would be something like this:

```vb
Dim numstudents As Integer
```

Using descriptive variable names takes a little more effort but can make your code easier to follow and less prone to error.
