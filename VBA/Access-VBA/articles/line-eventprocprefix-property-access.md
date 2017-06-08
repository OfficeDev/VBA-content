---
title: Line.EventProcPrefix Property (Access)
keywords: vbaac10.chm10326
f1_keywords:
- vbaac10.chm10326
ms.prod: access
api_name:
- Access.Line.EventProcPrefix
ms.assetid: d275d05d-5b38-d358-ebf1-3e3210afe704
ms.date: 06/08/2017
---


# Line.EventProcPrefix Property (Access)

Gets or sets the prefix portion of an event procedure name. Read/write  **String**.


## Syntax

 _expression_. **EventProcPrefix**

 _expression_ A variable that represents a **Line** object.


## Remarks

For example, if you have a command button with an event procedure named Details_Click, the  **EventProcPrefix** property returns the string "Details".

Microsoft Access adds the prefix portion of an event procedure name to the event name with an underscore character (_).


## See also


#### Concepts


[Line Object](line-object-access.md)

