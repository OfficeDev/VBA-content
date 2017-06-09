---
title: ObjectFrame.EventProcPrefix Property (Access)
keywords: vbaac10.chm11557
f1_keywords:
- vbaac10.chm11557
ms.prod: access
api_name:
- Access.ObjectFrame.EventProcPrefix
ms.assetid: a38ca887-8d70-eb89-a1ac-fd7308d17c0d
ms.date: 06/08/2017
---


# ObjectFrame.EventProcPrefix Property (Access)

Gets or sets the prefix portion of an event procedure name. Read/write  **String**.


## Syntax

 _expression_. **EventProcPrefix**

 _expression_ A variable that represents an **ObjectFrame** object.


## Remarks

For example, if you have a command button with an event procedure named Details_Click, the  **EventProcPrefix** property returns the string "Details".

Microsoft Access adds the prefix portion of an event procedure name to the event name with an underscore character (_).


## See also


#### Concepts


[ObjectFrame Object](objectframe-object-access.md)

