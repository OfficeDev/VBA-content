---
title: End Statement
keywords: vblr6.chm1008904
f1_keywords:
- vblr6.chm1008904
ms.prod: office
ms.assetid: 5cbb1c20-2afa-782e-52bb-7aafc604a927
ms.date: 06/08/2017
---


# End Statement

Ends a [procedure](vbe-glossary.md) or block.

 **Syntax**

 **End**

 **End** **Function**
 **End** **If**
 **End Property**
 **End Select**
 **End Sub**
 **End Type**
 **End With**
The  **End** statement syntax has these forms:


|**Statement**|**Description**|
|:-----|:-----|
|**End**|Terminates execution immediately. Never required by itself but may be placed anywhere in a procedure to end code execution, close files opened with the  **Open** statement and to clear[variables](vbe-glossary.md).|
|**End Function**|Required to end a  **Function** statement.|
|**End If**|Required to end a block  **If…Then…Else** statement.|
|**End Property**|Required to end a  **Property Let**, **Property Get**, or **Property Set** procedure.|
|**End Select**|Required to end a  **Select Case** statement.|
|**End Sub**|Required to end a  **Sub** statement.|
|**End Type**|Required to end a [user-defined type](vbe-glossary.md) definition ( **Type** statement).|
|**End With**|Required to end a  **With** statement.|
 **Remarks**
When executed, the  **End** statement resets all[module-level](vbe-glossary.md) variables and all static local variables in all[modules](vbe-glossary.md). To preserve the value of these variables, use the  **Stop** statement instead. You can then resume execution while preserving the value of those variables.

 **Note**  The  **End** statement stops code execution abruptly, without invoking the Unload, QueryUnload, or Terminate event, or any other Visual Basic code. Code you have placed in the Unload, QueryUnload, and Terminate events of[forms](vbe-glossary.md) and[class modules](vbe-glossary.md) is not executed. Objects created from class modules are destroyed, files opened using the **Open** statement are closed, and memory used by your program is freed. Object references held by other programs are invalidated.

The  **End** statement provides a way to force your program to halt. For normal termination of a Visual Basic program, you should unload all forms. Your program closes as soon as there are no other programs holding references to objects created from your public class modules and no code executing.

## Example

This example uses the  **End** Statement to end code execution if the user enters an invalid password.


```vb
Sub Form_Load 
 Dim Password, Pword 
 PassWord = "Swordfish" 
 Pword = InputBox("Type in your password") 
 If Pword <> PassWord Then 
 MsgBox "Sorry, incorrect password" 
 EndEnd IfEnd Sub
```


