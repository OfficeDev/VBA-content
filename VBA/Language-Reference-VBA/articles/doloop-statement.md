---
title: Do...Loop Statement
keywords: vblr6.chm1008790
f1_keywords:
- vblr6.chm1008790
ms.prod: office
ms.assetid: f1ac3901-238d-3e38-45dc-f659fd88c23b
ms.date: 06/08/2017
---


# Do...Loop Statement

Repeats a block of [statements](vbe-glossary.md) while a condition is **True** or until a condition becomes **True**.

 **Syntax**

 **Do** [{ **While** |**Until** } _condition_ ]
[ _statements_ ]
[ **Exit Do** ]
[ _statements_ ]

 **Loop**
Or, you can use this syntax:
 **Do**
[ _statements_ ]
[ **Exit Do** ]
[ _statements_ ]
 **Loop** [{ **While** |**Until** } _condition_ ]
The  **Do Loop** statement syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _condition_|Optional. [Numeric expression](vbe-glossary.md) or[string expression](vbe-glossary.md) that is **True** or **False**. If _condition_ is[Null](vbe-glossary.md),  _condition_ is treated as **False**.|
| _statements_|One or more statements that are repeated while, or until,  _condition_ is **True**.|
 **Remarks**
Any number of  **Exit Do** statements may be placed anywhere in the **Do…Loop** as an alternate way to exit a **Do…Loop**. **Exit Do** is often used after evaluating some condition, for example, **If…Then**, in which case the **Exit Do** statement transfers control to the statement immediately following the **Loop**.
When used within nested  **Do…Loop** statements, **Exit Do** transfers control to the loop that is one nested level above the loop where **Exit Do** occurs.

## Example

This example shows how  **Do...Loop** statements can be used. The inner **Do...Loop** statement loops 10 times, sets the value of the flag to **False**, and exits prematurely using the **Exit Do** statement. The outer loop exits immediately upon checking the value of the flag.


```vb
Dim Check, Counter 
Check = True: Counter = 0 ' Initialize variables. 
Do ' Outer loop. 
 Do While Counter < 20 ' Inner loop. 
 Counter = Counter + 1 ' Increment Counter. 
 If Counter = 10 Then ' If condition is True. 
 Check = False ' Set value of flag to False. 
 Exit Do ' Exit inner loop. 
 End If 
 LoopLoop Until Check = False ' Exit outer loop immediately. 

```


