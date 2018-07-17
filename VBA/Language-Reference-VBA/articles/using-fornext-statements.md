---
title: Using For...Next Statements (VBA)
keywords: vbcn6.chm1076682
f1_keywords:
- vbcn6.chm1076682
ms.prod: office
ms.assetid: fe6e66a7-a9d3-d363-65c5-00d35bb407bd
ms.date: 06/08/2017
---

# Using For...Next Statements

You can use  **For...Next** statements to repeat a block of [statements](vbe-glossary.md) a specific number of times. **For** loops use a counter [variable](vbe-glossary.md) whose value is increased or decreased with each repetition of the loop.

The following [procedure](vbe-glossary.md) makes the computer beep 50 times. The **For** statement specifies the counter variable and its start and end values. The **Next** statement increments the counter variable by 1.

```vb
Sub Beeps() 
    For x = 1 To 50 
        Beep 
    Next x 
End Sub
```

Using the  **Step** [keyword](vbe-glossary.md), you can increase or decrease the counter variable by the value you specify. In the following example, the counter variable  `j` is incremented by 2 each time the loop repeats. When the loop is finished, `total` is the sum of 2, 4, 6, 8, and 10.

```vb
Sub TwosTotal() 
    For j = 2 To 10 Step 2 
        total = total + j 
    Next j 
    MsgBox "The total is " &; total 
End Sub
```

To decrease the counter variable, use a negative  **Step** value. To decrease the counter variable, you must specify an end value that is less than the start value. In the following example, the counter variable `myNum` is decreased by 2 each time the loop repeats. When the loop is finished, `total` is the sum of 16, 14, 12, 10, 8, 6, 4, and 2.

```vb
Sub NewTotal() 
    For myNum = 16 To 2 Step -2 
        total = total + myNum 
    Next myNum 
    MsgBox "The total is " &; total 
End Sub
```

 **Note**  It's not necessary to include the counter variable name after the  **Next** statement. In the preceding examples, the counter variable name was included for readability.

You can exit a  **For...Next** statement before the counter reaches its end value by using the **Exit For** statement. For example, when an error occurs, use the **Exit For** statement in the **True** statement block of either an **If...Then...Else** statement or a **Select Case** statement that specifically checks for the error. If the error doesn't occur, then the **If…Then…Else** statement is **False**, and the loop will continue to run as expected.
