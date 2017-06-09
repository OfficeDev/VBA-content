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


|**Statement**|**Description**|
|:-----|:-----|
|**Exit Do**|Provides a way to exit a  **Do...Loop** statement. It can be used only inside a **Do...Loop** statement. **Exit Do** transfers control to the[statement](vbe-glossary.md) following the **Loop** statement. When used within nested **Do...Loop** statements, **Exit Do** transfers control to the loop that is one nested level above the loop where **Exit** **Do** occurs.|
|**Exit For**|Provides a way to exit a  **For** loop. It can be used only in a **For...Next** or **For** **Each...Next** loop. **Exit For** transfers control to the statement following the **Next** statement. When used within nested **For** loops, **Exit For** transfers control to the loop that is one nested level above the loop where **Exit For** occurs.|
|**Exit Function**|Immediately exits the  **Function**[procedure](vbe-glossary.md) in which it appears. Execution continues with the statement following the statement that called the **Function**.|
|**Exit Property**|Immediately exits the  **Property** procedure in which it appears. Execution continues with the statement following the statement that called the **Property** procedure.|
|**Exit Sub**|Immediately exits the  **Sub** procedure in which it appears. Execution continues with the statement following the statement that called the **Sub** procedure.|
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


