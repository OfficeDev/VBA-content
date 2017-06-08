---
title: Calling Sub and Function Procedures
keywords: vbcn6.chm1076673
f1_keywords:
- vbcn6.chm1076673
ms.prod: office
ms.assetid: 17a9dec1-d8f2-584c-324f-164b4f7b156f
ms.date: 06/08/2017
---


# Calling Sub and Function Procedures

To call a  **Sub** procedure from another [procedure](vbe-glossary.md), type the name of the procedure and include values for any required [arguments](vbe-glossary.md). The  **Call** statement is not required, but if you use it, you must enclose any arguments in parentheses.

You can use a  **Sub** procedure to organize other procedures so they are easier to understand and debug. In the following example, the **Sub** procedure `Main` calls the **Sub** procedure `MultiBeep`, passing the value 56 for its argument. After  `MultiBeep` runs, control returns to `Main`, and  `Main` calls the **Sub** procedure `Message`.  `Message` displays a message box; when the user clicks **OK**, control returns to `Main`, and calls the  **Sub** procedure `MultiBeep`, passing the value 56 for its argument. After  `MultiBeep` runs, control returns to `Main`, and  `Main` calls the **Sub** procedure `Message`.  `Message` displays a message box; when the user clicks **OK**, control returns to `Main`, and  `Main` finishes.



```vb
Sub Main() 
 MultiBeep 56 
 Message 
End Sub 
 
Sub MultiBeep(numbeeps) 
 For counter = 1 To numbeeps 
 Beep 
 Next counter 
End Sub 
 
Sub Message() 
 MsgBox "Time to take a break!" 
End Sub
```


## Calling Sub Procedures with More than One Argument

The following example shows two ways to call a  **Sub** procedure with more than one argument. The second time is called, parentheses are required around the arguments because the **Call** statement is used.


```vb
Sub Main() 
 HouseCalc 99800, 43100 
 Call HouseCalc(380950, 49500) 
End Sub 
 
Sub HouseCalc(price As Single, wage As Single) 
 If 2.5 * wage <= 0.8 * price Then 
 MsgBox "You cannot afford this house." 
 Else 
 MsgBox "This house is affordable." 
 End If 
End Sub
```


## Using Parentheses when Calling Function Procedures

To use the return value of a function, assign the function to a [variable](vbe-glossary.md) and enclose the arguments in parentheses, as shown in the following example.


```vb
Answer3 = MsgBox("Are you happy with your salary?", 4, "Question 3") 

```

If you're not interested in the return value of a function, you can call a function the same way you call a  **Sub** procedure. Omit the parentheses, list the arguments, and do not assign the function to a variable, as shown in the following example.




```vb
MsgBox "Task Completed!", 0, "Task Box" 

```

If you include parentheses in the preceding example, the statement causes a syntax error.


## Passing Named Arguments

A statement in a  **Sub** or **Function** procedure can pass values to called procedures using [named arguments](vbe-glossary.md). You can list named arguments in any order. A named argument consists of the name of the argument followed by a colon and an equal sign ( **:=** ), and the value assigned to the argument.

The following example calls the  **MsgBox** function using named arguments with no return value.




```vb
MsgBox Title:="Task Box", Prompt:="Task Completed!" 

```

The following example calls the  **MsgBox** function using named arguments. The return value is assigned to the variable .




```
answer3 = MsgBox(Title:="Question 3", _ 
Prompt:="Are you happy with your salary?", Buttons:=4) 

```


