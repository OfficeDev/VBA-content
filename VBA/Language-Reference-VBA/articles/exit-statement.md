---
title: Exit Statement
keywords: vblr6.chm1008916
f1_keywords:
- vblr6.chm1008916
ms.prod: office
ms.assetid: 2a1f4605-8220-c5b1-3760-c710f0535aa8
ms.date: 06/08/2017
---


# Exit Statement

Exits a block of  **Do…Loop**, **For…Next**, **Function**, **Sub**, or **Property** code.

 **Syntax**

 **Exit** **Do**

 **Exit For**
 **Exit Function**
 **Exit Property**
 **Exit Sub**
The  **Exit** statement syntax has these forms:


| <strong>Statement</strong>     | <strong>Description</strong>                                                                                                                                                                                                                                                                                                                                                                                                                                                          |
|:-------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <strong>Exit Do</strong>       | Provides a way to exit a  <strong>Do...Loop</strong> statement. It can be used only inside a <strong>Do...Loop</strong> statement. <strong>Exit Do</strong> transfers control to the[statement](vbe-glossary.md) following the <strong>Loop</strong> statement. When used within nested <strong>Do...Loop</strong> statements, <strong>Exit Do</strong> transfers control to the loop that is one nested level above the loop where <strong>Exit</strong> <strong>Do</strong> occurs. |
| <strong>Exit For</strong>      | Provides a way to exit a  <strong>For</strong> loop. It can be used only in a <strong>For...Next</strong> or <strong>For</strong> <strong>Each...Next</strong> loop. <strong>Exit For</strong> transfers control to the statement following the <strong>Next</strong> statement. When used within nested <strong>For</strong> loops, <strong>Exit For</strong> transfers control to the loop that is one nested level above the loop where <strong>Exit For</strong> occurs.          |
| <strong>Exit Function</strong> | Immediately exits the  <strong>Function</strong>[procedure](vbe-glossary.md) in which it appears. Execution continues with the statement following the statement that called the <strong>Function</strong>.                                                                                                                                                                                                                                                                           |
| <strong>Exit Property</strong> | Immediately exits the  <strong>Property</strong> procedure in which it appears. Execution continues with the statement following the statement that called the <strong>Property</strong> procedure.                                                                                                                                                                                                                                                                                   |
| <strong>Exit Sub</strong>      | Immediately exits the  <strong>Sub</strong> procedure in which it appears. Execution continues with the statement following the statement that called the <strong>Sub</strong> procedure.                                                                                                                                                                                                                                                                                             |

 **Remarks**
Do not confuse  **Exit** statements with **End** statements. **Exit** does not define the end of a structure.

## Example

This example uses the  **Exit** statement to exit a **For...Next** loop, a **Do...Loop**, and a **Sub** procedure.


```vb
Sub ExitStatementDemo() 
Dim I, MyNum 
 Do ' Set up infinite loop. 
 For I = 1 To 1000 ' Loop 1000 times. 
 MyNum = Int(Rnd * 1000) ' Generate random numbers. 
 Select Case MyNum ' Evaluate random number. 
 Case 7: Exit For ' If 7, exit For...Next. 
 Case 29: Exit Do ' If 29, exit Do...Loop. 
 Case 54: Exit Sub ' If 54, exit Sub procedure. 
 End Select 
 Next I 
 Loop 
End Sub
```


