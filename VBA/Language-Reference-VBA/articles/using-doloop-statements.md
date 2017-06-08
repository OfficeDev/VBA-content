---
title: Using Do...Loop Statements
keywords: vbcn6.chm1076681
f1_keywords:
- vbcn6.chm1076681
ms.prod: office
ms.assetid: aa3322b6-80a6-d3c6-86b7-4ea6151f0616
ms.date: 06/08/2017
---


# Using Do...Loop Statements

You can use  **Do...Loop** statements to run a block of[statements](vbe-glossary.md) an indefinite number of times. The statements are repeated either while a condition is **True** or until a condition becomes **True**.


## Repeating Statements While a Condition is True

There are two ways to use the  **While**[keyword](vbe-glossary.md) to check a condition in a **Do...Loop** statement. You can check the condition before you enter the loop , or you can check it after the loop has run at least once.

In the following  `ChkFirstWhile` procedure, you check the condition before you enter the loop. If `myNum` is set to 9 instead of 20, the statements inside the loop will never run. In the `ChkLastWhile` procedure, the statements inside the loop run only once before the condition becomes **False**.




```vb
Sub ChkFirstWhile() 
    counter = 0 
    myNum = 20 
    Do While myNum > 10 
        myNum = myNum - 1 
        counter = counter + 1 
    Loop 
    MsgBox "The loop made " &; counter &; " repetitions." 
End Sub 
 
Sub ChkLastWhile() 
    counter = 0 
    myNum = 9 
    Do 
        myNum = myNum - 1 
        counter = counter + 1 
    Loop While myNum > 10 
    MsgBox "The loop made " &; counter &; " repetitions." 
End Sub
```


## Repeating Statements Until a Condition Becomes True

There are two ways to use the  **Until** keyword to check a condition in a **Do...Loop** statement. You can check the condition before you enter the loop (as shown in the `ChkFirstUntil` procedure), or you can check it after the loop has run at least once (as shown in the `ChkLastUntil` procedure). Looping continues while the condition remains **False**.


```vb
Sub ChkFirstUntil() 
    counter = 0 
    myNum = 20 
    Do Until myNum = 10 
        myNum = myNum - 1 
        counter = counter + 1 
    Loop 
    MsgBox "The loop made " &; counter &; " repetitions." 
End Sub 
 
Sub ChkLastUntil() 
    counter = 0 
    myNum = 1 
    Do 
        myNum = myNum + 1 
        counter = counter + 1 
    Loop Until myNum = 10 
    MsgBox "The loop made " &; counter &; " repetitions." 
End Sub
```


## Exiting a Do...Loop Statement from Inside the Loop

You can exit a  **Do...Loop** using the **Exit Do** statement. For example, to exit an endless loop, use the **Exit Do** statement in the **True** statement block of either an **If...Then...Else** statement or a **Select Case** statement. If the condition is **False**, the loop will run as usual.

In the following example,  `myNum` is assigned a value that creates an endless loop. The **If...Then...Else** statement checks for this condition, and then exits, preventing endless looping.




```vb
Sub ExitExample() 
    counter = 0 
    myNum = 9 
    Do Until myNum = 10 
        myNum = myNum - 1 
        counter = counter + 1 
        If myNum < 10 Then Exit Do 
    Loop 
    MsgBox "The loop made " &; counter &; " repetitions." 
End Sub
```


 **Note**  To stop an endless loop, press ESC or CTRL+BREAK.


