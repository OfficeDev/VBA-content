---
title: Stop Statement
keywords: vblr6.chm1009033
f1_keywords:
- vblr6.chm1009033
ms.prod: office
ms.assetid: 9b6b5394-9b19-8f18-216c-ac64b165218f
ms.date: 06/08/2017
---


# Stop Statement

Suspends execution.

 **Syntax**

 **Stop**

 **Remarks**
You can place  **Stop** statements anywhere in[procedures](vbe-glossary.md) to suspend execution. Using the **Stop** statement is similar to setting a[breakpoint](vbe-glossary.md) in the code.
The  **Stop** statement suspends execution, but unlike **End**, it doesn't close any files or clear[variables](vbe-glossary.md), unless it is in a compiled executable (.exe) file.

## Example

This example uses the  **Stop** statement to suspend execution for each iteration through the **For...Next** loop.


```vb
Dim I 
For I = 1 To 10 ' Start For...Next loop. 
 Debug.Print I ' Print I to the Immediate window. 
 Stop ' Stop during each iteration. 
Next I 

```


