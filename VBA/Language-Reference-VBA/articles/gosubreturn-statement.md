---
title: GoSub...Return Statement
keywords: vblr6.chm1008934
f1_keywords:
- vblr6.chm1008934
ms.prod: office
ms.assetid: 5aafb93f-0baf-f319-d8dd-96a14095d62d
ms.date: 06/08/2017
---


# GoSub...Return Statement

Branches to and returns from a subroutine within a [procedure](vbe-glossary.md).

 **Syntax**

 **GoSub**_line_
 `...`
 _line_

 _line_
 `...`

 **Return**
The  _line_[argument](vbe-glossary.md) can be any[line label](vbe-glossary.md) or[line number](vbe-glossary.md).
 **Remarks**
You can use  **GoSub** and **Return** anywhere in a procedure, but **GoSub** and the corresponding **Return** statement must be in the same procedure. A subroutine can contain more than one **Return** statement, but the first **Return** statement encountered causes the flow of execution to branch back to the[statement](vbe-glossary.md) immediately following the most recently executed **GoSub** statement.

 **Note**  You can't enter or exit  **Sub** procedures with **GoSub...Return**.


 **Tip**  Creating separate procedures that you can call may provide a more structured alternative to using  **GoSub...Return**.


## Example

This example uses  **GoSub** to call a subroutine within a **Sub** procedure. The **Return** statement causes the execution to resume at the statement immediately following the **GoSub** statement. The **Exit Sub** statement is used to prevent control from accidentally flowing into the subroutine.


```vb
Sub GosubDemo() 
Dim Num 
' Solicit a number from the user. 
 Num = InputBox("Enter a positive number to be divided by 2.") 
' Only use routine if user enters a positive number. 
 If Num > 0 Then GoSub MyRoutine 
 Debug.Print Num 
 Exit Sub ' Use Exit to prevent an error. 
MyRoutine: 
 Num = Num/2 ' Perform the division. 
 Return ' Return control to statement. 
End Sub ' following the GoSub statement. 

```


