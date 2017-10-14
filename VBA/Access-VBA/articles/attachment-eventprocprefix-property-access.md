---
title: Attachment.EventProcPrefix Property (Access)
keywords: vbaac10.chm13912
f1_keywords:
- vbaac10.chm13912
ms.prod: access
api_name:
- Access.Attachment.EventProcPrefix
ms.assetid: f58670ff-b42c-69eb-0561-90ce5cc40d19
ms.date: 06/08/2017
---


# Attachment.EventProcPrefix Property (Access)

Gets or sets the prefix portion of an event procedure name. Read/write  **String**.


## Syntax

 _expression_. **EventProcPrefix**

 _expression_ A variable that represents an **Attachment** object.


## Remarks

For example, if you have a command button with an event procedure named Details_Click, the  **EventProcPrefix** property returns the string "Details".

Microsoft Access adds the prefix portion of an event procedure name to the event name with an underscore character (_).


## See also


#### Concepts


[Attachment Object](attachment-object-access.md)

