---
title: Can't ReDim, Erase, or assign to Variant that contains array whose element is With object
keywords: vblr6.chm1040376
f1_keywords:
- vblr6.chm1040376
ms.prod: office
ms.assetid: 0fe8d19c-00f3-ceb9-5ce9-fc349221de6a
ms.date: 06/08/2017
---


# Can't ReDim, Erase, or assign to Variant that contains array whose element is With object

This error has the following causes and solutions:



- You've attempted to  **ReDim**, **Erase**, or assign to a **Variant** a variable whose element is a With object. For example, the following code produces this error:
    
```vb
Type Test
   Name as Integer
End Type

Sub Main()
   Dim c(0) As Test
   Dim e e = c
   With e(0)
      ReDim e(1)
   End With
End Sub
  ```


For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

