# Working with Cells and Ranges

When writing VB code, you can use variables, for loops, and all other VB types and statements. However, most of your code will be dealing with values stored in cells and ranges.

## The Range Object

A **range** is set of cells. It can be one cell or multiple cells. A range is an object within a worksheet object. For example, the following statement sets the value of cell C23 to a formula referencing a set of named cells:

```vb
Worksheets("trapchan").Range("C23").Value = "=(u/n)*A*Rh^(2/3)*S^(1/2)"
```

If the "trapchan" worksheet is the active sheet, the first part can be left off as follows:

```vb
Range("C23").Value = "=(u/n)*A*Rh^(2/3)*S^(1/2)"
```

The **.Value** part is optional. You can also write:

```vb
Range("C23") = "=(u/n)*A*Rh^(2/3)*S^(1/2)"
```

You have to remember to put the double quotes around the cell address. If the cell C23 has been named "Q" in the spreadsheet, you can reference the range as follows:

```vb
Range("Q").Value = "=(u/n)*A*Rh^(2/3)*S^(1/2)"
```

To get something from a cell and put it in a variable, you just do things in reverse:

```vb
x = Range("B14").Value
```

## The Cells Object

Another way to interact with a cell is to use the **Cells(rowindex, columnindex)** function. For example:

```vb
Cells(2, 5).Value = "=(u/n)*A*Rh^(2/3)*S^(1/2)"
```

or

```vb
x = val(Cells(2, 14).Value)
```

Once again, the value part is optional since it is the default property of both the cell and range object.

## Working with Multiple Cells

A range can also encompass a set of cells. The following code selects a block of cells:

```vb
Range("A1:C5").Select
```

or

```vb
Range("A1", "C5").Select
```

In some cases, it is useful to reference a range of cells using integers representing the row and column of the cell. This can be accomplished with the **Cells** object. For example:

```vb
Range(Cells(1,1), Cells(3,5)).Select
```

The problem with referring to specific cells in your code is that if you change the location of data on your spreadsheet, you need to go through your code and make sure all of the addresses are updated. In some cases it is useful to define the ranges you are dealing with using global constants at the top of your VB code. Then when you reference a range, you can use the constants. If the range ever changes, you only need to update your code in one location. For example:

```vb
Const TableRange As String = "A4:D50"
Const NameRange As String = "A4:D50"
Const ScoreRange As String = "D4:D50"

.
.
Range(TableRange).ClearContents
.
.
```

An even better approach is to get into the habit of naming cells and ranges on your spreadsheet. Then your VB code can always refer to ranges by names. Then, if you change the location or domain of a named range, you generally don't need to update your VB code. For example:

```vb
Range("TableRange").ClearContents
Range("NameRange").ClearContents
Range("ScoreRange").ClearContents
```

## Looping Through Cells

One of the most common things we do with VB code is to traverse or loop through a set of cells. There are several ways this can be accomplished. One way is to use the **Cells** object. The following code loops through a table of cells located in the range B4:F20:

```vb
Dim row As Integer
Dim col As Integer

For row = 4 To 20
   For col = 2 To 6
      If Cells(row, col) = "" Then
         Cells(row, col).Interior.Color = vbRed
      End If
   Next col
Next row
```

In most cases, it doesn't matter what order the cells are traversed, as long as each cell is visited. In such cases the **For Each ... Next** looping style may be used. The following code does the same thing as the nested loop shown above:

```vb
Dim cell As Variant

For Each cell In Range("B4:F20")
   If cell = "" Then
      cell.Interior.Color = vbRed
   End If
Next cell
```

Another option is to create your own range objects. A range object is essentially a range type variable. The following code defines three ranges:

```vb
Dim coordrange As Range
Dim xrange As Range
Dim yrange As Range

Set xrange = Range(Cells(3, 1), Cells(100, 1))
Set yrange = Range(Cells(3, 2), Cells(100, 2))
Set coordrange = Range(Cells(3, 1), Cells(100, 2))
```

Once a set of range objects has been defined, you can easily manipulate the cells in the range object. For example, the following code clears the contents of all the cells in the coordrange object:

```vb
coordrange.Clear
```

Once again, these ranges can be traversed using the **For Each ... Next** syntax.

```vb
Dim cell As Variant

For Each cell In xrange
    cell.Value = "0.0"
Next cell
```

One of the most useful things you can do with VBA in Excel is to allow the user to enter a list of numbers where the size of the list can vary. The following code searches through a list and copies the numbers in the list into an array. It stops copying the numbers when it reaches a blank cell.

```vb
'Get the x values
i = 0
For Each cell In Range("B4:B23")
    If cell.Value = "" Then
        numpts = i
        Exit For
    Else
        i = i + 1
        x(i) = cell.Value
    End If
Next cell
```

## R1C1 Notation

Another case where we deal with ranges in VB code is when we record macros and then modify the macros to make custom subs. If we create a formula while recording the macro, the formula will often be recorded using what is called "R1C1 notation". For example, here is a snippet of code from a recorded macro:

```vb
Range("F11").Select
ActiveCell.FormulaR1C1 = "=SUM(R[-4]C:R[-2]C)"
```

This is the same as:

```vb
Range("F11").Select
ActiveCell.Formula = "=SUM(F7:F9)"
```

So what is the advantage of the R1C1 notation? If you need to modify the formula using variables or some other run-time conditions, the R1C1 notation can be easier to manipulate. The R1C1 notation makes it possible to refer to both rows and columns using integers rather than letters and it makes it very easy to indicate relative vs. absolute references. The following table illustrates how the R1C1 notation works:

| R1C1 Expression | Equivalent Expression* |
|-----------------|------------------------|
| `R[-4]C:R[-2]C` | F7:F9 |
| `R[+2]C:R[+4]C` | F13:F15 |
| `R4C2:RC[-1]` | $B$4:E11 |
| `RC:R15C10` | F11:$J$15 |

*Assumes reference cell = F11

In other words, R10 means "an absolute reference to row 10" whereas R[-2] means "a relative reference to two rows above the cell in question" and R (without a number) means "on the same row as the cell in question".

## Exercises

You may wish to complete following exercises to gain practice with and reinforce the topics covered in this chapter:

<div class="exercise-grid" data-columns="4">
<div class="exercise-header">Description</div>
<div class="exercise-header">Difficulty</div>
<div class="exercise-header">Start</div>
<div class="exercise-header">Solution</div>
<div class="exercise-cell"><strong>Range Basics -</strong> Run through 5 quick exercises that use ranges.</div>
<div class="exercise-cell">Easy</div>
<div class="exercise-cell"><a href="files/range_basics.xlsm">range_basics.xlsm</a></div>
<div class="exercise-cell"><a href="files/range_basics_key.xlsm">range_basics_key.xlsm</a></div>
<div class="exercise-cell"><strong>Calculator -</strong> Calculate the given equations from three variables.</div>
<div class="exercise-cell">Medium</div>
<div class="exercise-cell"><a href="files/calculator.xlsm">calculator.xlsm</a></div>
<div class="exercise-cell"><a href="files/calculator_key.xlsm">calculator_key.xlsm</a></div>
<div class="exercise-cell"><strong>Sum -</strong> Use ranges to calculate a sum and an adjusted sum of the tabulated data.</div>
<div class="exercise-cell">Medium</div>
<div class="exercise-cell"><a href="files/sum.xlsm">sum.xlsm</a></div>
<div class="exercise-cell"><a href="files/sum_key.xlsm">sum_key.xlsm</a></div>
<div class="exercise-cell"><strong>Empty Cells -</strong> Calculate the number of empty cells from the range of data. Also change the range object properties to display a red cell for an empty cell.</div>
<div class="exercise-cell">Hard</div>
<div class="exercise-cell"><a href="files/empty_cells.xlsm">empty_cells.xlsm</a></div>
<div class="exercise-cell"><a href="files/empty_cells_key.xlsm">empty_cells_key.xlsm</a></div>
</div>