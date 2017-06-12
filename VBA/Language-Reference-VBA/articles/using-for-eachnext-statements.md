---
title: Using For Each...Next Statements
keywords: vbcn6.chm1076683
f1_keywords:
- vbcn6.chm1076683
ms.prod: office
ms.assetid: 76df8944-219a-c28b-c449-39a3108c11be
ms.date: 06/08/2017
---


# Using For Each...Next Statements

 **For Each...Next** statements repeat a block of[statements](vbe-glossary.md) for each[object](vbe-glossary.md) in a[collection](vbe-glossary.md) or each element in an[array](vbe-glossary.md). Visual Basic automatically sets a [variable](vbe-glossary.md) each time the loop runs. For example, the following[procedure](vbe-glossary.md) closes all forms except the form containing the procedure that's running.


```vb
Sub CloseForms() 
 For Each frm In Application.Forms 
 If frm.Caption <> Screen. ActiveForm.Caption Then frm.Close 
 Next 
End Sub
```


The following code loops through each element in an array and sets the value of each to the value of the index variable I.




```vb
Dim TestArray(10) As Integer, I As Variant 
For Each I In TestArray 
 TestArray(I) = I 
Next I 

```


## Looping Through a Range of Cells

Use a  **For Each...Next** loop to loop through the cells in a range. The following procedure loops through the range A1:D10 on Sheet1 and sets any number whose absolute value is less than 0.01 to 0 (zero).


```vb
Sub RoundToZero() 
 For Each myObject in myCollection 
 If Abs(myObject.Value) < 0.01 Then myObject.Value = 0 
 Next 
End Sub
```


## Exiting a For Each...Next Loop Before it is Finished

You can exit a  **For Each...Next** loop using the **Exit For** statement. For example, when an error occurs, use the **Exit For** statement in the **True** statement block of either an **If...Then...Else** statement or a **Select Case** statement that specifically checks for the error. If the error does not occur, then the **If…Then…Else** statement is **False** and the loop continues to run as expected.

The following example tests for the first cell in the range A1:B5 that does not contain a number. If such a cell is found, a message is displayed and  **Exit For** exits the loop.




```vb
Sub TestForNumbers() 
 For Each myObject In MyCollection 
 If IsNumeric(myObject.Value) = False Then 
 MsgBox "Object contains a non-numeric value." 
 Exit For 
 End If 
 Next c 
End Sub
```


