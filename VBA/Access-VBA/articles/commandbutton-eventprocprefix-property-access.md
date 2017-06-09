---
title: CommandButton.EventProcPrefix Property (Access)
keywords: vbaac10.chm10447
f1_keywords:
- vbaac10.chm10447
ms.prod: access
api_name:
- Access.CommandButton.EventProcPrefix
ms.assetid: 89611b46-0c56-d855-9e4d-d1a301f300ae
ms.date: 06/08/2017
---


# CommandButton.EventProcPrefix Property (Access)

Gets or sets the prefix portion of an event procedure name. Read/write  **String**.


## Syntax

 _expression_. **EventProcPrefix**

 _expression_ A variable that represents a **CommandButton** object.


## Remarks

For example, if you have a command button with an event procedure named Details_Click, the  **EventProcPrefix** property returns the string "Details".

Microsoft Access adds the prefix portion of an event procedure name to the event name with an underscore character (_).


## See also


#### Concepts


[CommandButton Object](commandbutton-object-access.md)

