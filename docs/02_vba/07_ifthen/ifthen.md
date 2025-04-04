# Decisions and Conditions - Writing If Statements

There are countless ocassions when writing code where we need to execute code only if certain conditions are met. This can be accomplished by writing If statements.

## Syntax

The general syntax of an If statement is as follows:

```vbnet
If condition_1 Then
    statement(s)
ElseIf condition_2 Then 
    statement(s)

...

ElseIf condition_n Then
    statement(s)
Else
    statement(s)
End If
```

The If ... Then and the End If parts are required. All of the other parts are optional and are only used when needed. For example, in some cases you don't need any of the Else options:

```vbnet
If Range("B5") = "" Then
    MsgBox "Error: Cell B5 cannot be empty"
    Exit Sub
End If
```

Note that you can put as many statements as you want between the If and End If lines. Each of the statements is executed if the condition is true. You do not have to indent, but it is STRONGLY recommended as it makes your code much easier to follow.

In some cases you may wish to include an Else clause that is executed when the condition is false.

```vbnet
If x = 0 Then
    MsgBox "Error: Cannot divide by zero"
Else
    y = 1 / x
End If
```

In other cases, you may need to check several conditions, each of which is mutually exclusive.

```vbnet
If yourteam = "BYU" Then
    MsgBox "You are cool"
ElseIf yourteam = "Utah" Then
    MsgBox "You are NOT cool"
ElseIf yourteam = "Utah State" Then
    MsgBox "What is an aggie?"
Else
    MsgBox "I do not care"
End If
```

Each of the conditions is checked in sequence starting at the top. Once a condition is found that evaluates to True, none of the remaining conditions are tested and the flow of control exits the IF statement and jumps to the code immediately following the End If statement.

For the cases with no ElseIf clauses and simple one-line results, you can put your entire statement on a single line:

```vbnet
If yourteam = "BYU" Then MsgBox "You are cool"
```

or

```vbnet
If x = 0 Then y = 0 Else y = 1 / x
```
However, this is not recommended as it makes the code much harder to read. It is better to use the multi-line format shown above.

## Conditional Expressions

Every If statement requires at least one conditional expression. A conditional expression is an expression that returns either True or False when evaluated. Conditional expressions are generally formulated using a binary conditional operator. A binary operator takes two arguments, one on each side of the operator. Here is a list of the commonly used operators:

| Operator | Symbol | Example |
|----------|--------|---------|
| Equal | = | a = b |
| Not equal | <> | a <> b |
| Less than | < | x < y |
| Greater than | > | p > q |
| Less than Or equal to | <= | x <= 5.5 |
| Greater than or equal to | >= | y >= p |

Multiple conditional expressions can be combined with the And and Or operators. With the And operator, the combined expression is true if both conditions are true. With the Or operator, the combined expression is true if either of the two conditions is true. For example,

```vbnet
If myteam = "BYU" And yourteam = "BYU" Then
    MsgBox "High five!"
ElseIf yourteam = "Utah" Or yourteam = "USU" Then
    MsgBox "Boo!"
Else
    MsgBox "Nice to meet you"
End If
```

You may wish to combine more than two conditional expressions. In this case, it helps to use parentheses.

```vbnet
If (myteam = "BYU") And (yourteam = "Utah" Or yourteam = "USU") Then
    MsgBox "We are going to have a problem."
Else
    MsgBox "Nice to meet you"
End If
```

You can also use the Not operator to negate a statement. It is a unary operator and it negates the conditional expression that follows it. For example,

```vbnet
If (myteam = "BYU") And Not (yourteam = "Utah" Or yourteam = "USU") Then
    MsgBox "We are going to get along OK."
End If
```

Notice that Not True --> False and Not False --> True.

## Evaluating Number Ranges

When doing computations, it is common to need to determine if a number is inside a range. For example, in mathematics it is common two write a statement like this:

```
0 ≤ x ≤ 5
```

When writing this as a compound conditional expression, it is tempting to write it as follows:

```vbnet
If 0 <= x <= 5 Then
```

However, this is NOT logically equivalent to the statement shown above. For example, suppose that x = -10, which is outside the range and should make the expression evalue to False. The expression is evaluated in two parts from left to right, so the first part evaluated is 0 <= x, which returns a value of False. The value of False is then substituted for the first part of the expression and the remaining expression is then evaluated as False <= 5. Whenever a boolean value (True/False) is compared to numerical value, True = 1 and False = 0. Therefore, this expression is evaluated as 0 <= 5, which is True, leading to an incorrect result.

The proper way to write this expression in VB is:

```vbnet
If 0 <= x And x <= 5 Then
```

In this case, the two sides are evaluated independently and then combined with the And operator, resulting in the correct answer.

##If Statements and Controls

If statements are commonly used to determine the state of controls. Suppose you have a checkbox called chkResizeImage. You could check the state of the Value property as follows:

```vbnet
If chkResizeImage.Value = True Then
```

Note that for a checkbox and option control, the Value property is a Boolean variable that equals True if the control is selected, and False otherwise. Since the Value property is the default property for each of these objects, you can simplify this statement by omitting the .Value part as follows:

```vbnet
If chkResizeImage = True Then
```

Furthermore, this statement can be further simplified as follows:

```vbnet
If chkResizeImage Then
```

In other words, the = True part is redundant because chkResizeImage = True is logically equivalent to the value of chkResizeImage (the expression is true when chkResizeImage is true).

## Exercises

You may wish to complete following exercises to gain practice with and reinforce the topics covered in this chapter:

| Name  | Description	                                                                                                                           | Difficulty	| Start	                                        | Solution	|
|--------------|----------------------------------------------------------------------------------------------------------------------------------------|-------------|-----------------------------------------------|-----------|
| Score Keeper | Use an IF THEN expression to provide<br> feedback on a calculated score.	                                                              | Easy	                                                                                                                                  | [score_keeper.xlsm](files/score_keeper.xlsm)	 | [score_keeper_key.xlsm](files/score_keeper_key.xlsm)	|
| Density | Use an IF THEN expression with a check box<br> and a conditional expression to display <br>the correct density with the values given.	 | Medium	                                                                    | [density.xlsm](files/density.xlsm)	           | [density_key.xlsm](files/density_key.xlsm)	|
| Receipt |  Create an IF THEN expression that takes a<br> few conditions into consideration when <br>generating a receipt for a customer.	        | Hard	                                                                                                                                   |                 [receipt.xlsm](files/receipt.xlsm)                 | [receipt_key.xlsm](files/receipt_key.xlsm)	|
