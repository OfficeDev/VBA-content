---
title: ComboBox.EventProcPrefix Property (Access)
keywords: vbaac10.chm11371
f1_keywords:
- vbaac10.chm11371
ms.prod: access
api_name:
- Access.ComboBox.EventProcPrefix
ms.assetid: 79af6ac6-8876-ff72-16a8-5ec81ab6a0f8
ms.date: 06/08/2017
---


# ComboBox.EventProcPrefix Property (Access)

Gets or sets the prefix portion of an event procedure name. Read/write  **String**.


## Syntax

 _expression_. **EventProcPrefix**

 _expression_ A variable that represents a **ComboBox** object.


## Remarks

For example, if you have a command button with an event procedure named Details_Click, the  **EventProcPrefix** property returns the string "Details".

Microsoft Access adds the prefix portion of an event procedure name to the event name with an underscore character (_).


## See also


#### Concepts


[ComboBox Object](combobox-object-access.md)

