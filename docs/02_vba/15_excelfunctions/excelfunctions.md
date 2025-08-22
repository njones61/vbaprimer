# Calling Excel Functions from VB Code

One of the nice things about writing VB code inside Excel is that you can combine all of the power and flexibility of Visual Basic with the many tools and options in Excel. One of the best examples of this is that you can take advantage of all of the standard Excel worksheet functions inside your VB code. Calling an Excel worksheet function is simple. The Excel functions are available as methods within the **WorksheetFunction** object. You simply invoke the method and pass the arguments that the function requires (typically a range).

For example, if we were writing a simple formula to put in a cell to find the minimum value in a range of cells, we would write the following:

```vb
=Min(B4:F30)
```

The following code uses the same **Min** function, but invokes the function using VB code. The min value is stored in a variable called **minval**:

```vb
Dim minval As Double
minval = Application.WorksheetFunction.Min(Range("B4:F30"))
```

Notice the difference in how the range is specified. In the VB code, the range is specified as a range object.

The Application. portion is actually optional and can be omitted in most cases. Thus, the following code achieves the same thing:

```vb
Dim minval As Double
minval = WorksheetFunction.Min(Range("B4:F30"))
```

Here are some more examples:

```vb
Range("e5") = WorksheetFunction.sum(Range("b5:b29"))

'This is useful since VB does not have an inverse sin function
Dim x As Double
x = WorksheetFunction.Asin(0.223)

Dim i As Integer
i = 5
Range("H4") = WorksheetFunction.Fact(i)
```

## Exercises

You may wish to complete following exercises to gain practice with and reinforce the topics covered in this chapter:

<div class="exercise-grid" data-columns="4">
<div class="exercise-header">Description</div>
<div class="exercise-header">Difficulty</div>
<div class="exercise-header">Start</div>
<div class="exercise-header">Solution</div>
<div class="exercise-cell"><strong>Harmonic Mean -</strong> Use an Excel function within a custom function to calculate the harmonic mean from the tabulated data.</div>
<div class="exercise-cell">Easy</div>
<div class="exercise-cell"><a href="files/harmonic_mean.xlsm">harmonic_mean.xlsm</a></div>
<div class="exercise-cell"><a href="files/harmonic_mean_key.xlsm">harmonic_mean_key.xlsm</a></div>
<div class="exercise-cell"><strong>Law of Cosines -</strong> Calculate the Law of Cosines using an Excel function for Cosine within your sub.</div>
<div class="exercise-cell">Medium</div>
<div class="exercise-cell"><a href="files/law_of_cosines.xlsm">law_of_cosines.xlsm</a></div>
<div class="exercise-cell"><a href="files/law_of_cosines_key.xlsm">law_of_cosines_key.xlsm</a></div>
<div class="exercise-cell"><strong>Bill Payment -</strong> Use the APR Excel function within a custom function to calculate the number of months required to pay off a credit card bill.</div>
<div class="exercise-cell">Hard</div>
<div class="exercise-cell"><a href="files/bill_payments.xlsm">bill_payments.xlsm</a></div>
<div class="exercise-cell"><a href="files/bill_payments_key.xlsm">bill_payments_key.xlsm</a></div>
</div>