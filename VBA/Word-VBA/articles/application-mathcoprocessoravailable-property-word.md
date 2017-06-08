---
title: Application.MathCoprocessorAvailable Property (Word)
keywords: vbawd10.chm158335012
f1_keywords:
- vbawd10.chm158335012
ms.prod: word
api_name:
- Word.Application.MathCoprocessorAvailable
ms.assetid: 207b7f3f-4113-7069-51e3-10658ec3654f
ms.date: 06/08/2017
---


# Application.MathCoprocessorAvailable Property (Word)

 **True** if a math coprocessor is installed and available to Microsoft Word. Read-only **Boolean** .


## Syntax

 _expression_ . **MathCoprocessorAvailable**

 _expression_ An expression that returns an **[Application](application-object-word.md)** object.


## Example

This example displays a message indicating whether a math coprocessor is installed and available to Word.


```vb
If Application.MathCoprocessorAvailable = True Then 
 MsgBox "A math coprocessor is available." 
Else 
 MsgBox "A math coprocessor is not installed." 
End If
```


## See also


#### Concepts


[Application Object](application-object-word.md)

