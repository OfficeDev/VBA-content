---
title: DoCmd.Minimize Method (Access)
keywords: vbaac10.chm4157
f1_keywords:
- vbaac10.chm4157
ms.prod: access
api_name:
- Access.DoCmd.Minimize
ms.assetid: fa29ccaa-9d61-c5c3-fc32-f53a5d96ff05
ms.date: 06/08/2017
---


# DoCmd.Minimize Method (Access)

The  **Minimize** method carries out the Minimize action in Visual Basic.


## Syntax

 _expression_. **Minimize**

 _expression_ A variable that represents a **DoCmd** object.


## Remarks

You can use this method to remove a window from the screen while leaving the object open. You can also use this method to open an object without displaying its window. To display the object, use the  **SelectObject** method with either the **Maximize** or **Restore** method. The **Restore** method restores a minimized window to its previous size.

This method cannot be applied to module windows in the Visual Basic Editor (VBE). For information about how to affect module windows see the  **WindowState** property topic.


## See also


#### Concepts


[DoCmd Object](docmd-object-access.md)

