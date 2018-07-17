---
title: ListBox.EventProcPrefix Property (Access)
keywords: vbaac10.chm11218
f1_keywords:
- vbaac10.chm11218
ms.prod: access
api_name:
- Access.ListBox.EventProcPrefix
ms.assetid: 28f4d70b-8206-2481-9b83-c1bbc2767b82
ms.date: 06/08/2017
---


# ListBox.EventProcPrefix Property (Access)

Gets or sets the prefix portion of an event procedure name. Read/write  **String**.


## Syntax

 _expression_. **EventProcPrefix**

 _expression_ A variable that represents a **ListBox** object.


## Remarks

For example, if you have a command button with an event procedure named Details_Click, the  **EventProcPrefix** property returns the string "Details".

Microsoft Access adds the prefix portion of an event procedure name to the event name with an underscore character (_).


## See also


#### Concepts


[ListBox Object](listbox-object-access.md)

