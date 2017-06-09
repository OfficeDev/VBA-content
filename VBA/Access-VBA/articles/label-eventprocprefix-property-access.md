---
title: Label.EventProcPrefix Property (Access)
keywords: vbaac10.chm10188
f1_keywords:
- vbaac10.chm10188
ms.prod: access
api_name:
- Access.Label.EventProcPrefix
ms.assetid: 089ac12e-6ad3-4c0f-1025-be4c21f036c6
ms.date: 06/08/2017
---


# Label.EventProcPrefix Property (Access)

Gets or sets the prefix portion of an event procedure name. Read/write  **String**.


## Syntax

 _expression_. **EventProcPrefix**

 _expression_ A variable that represents a **Label** object.


## Remarks

For example, if you have a command button with an event procedure named Details_Click, the  **EventProcPrefix** property returns the string "Details".

Microsoft Access adds the prefix portion of an event procedure name to the event name with an underscore character (_).


## See also


#### Concepts


[Label Object](label-object-access.md)

