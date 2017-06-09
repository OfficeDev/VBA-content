---
title: BoundObjectFrame.EventProcPrefix Property (Access)
keywords: vbaac10.chm10907
f1_keywords:
- vbaac10.chm10907
ms.prod: access
api_name:
- Access.BoundObjectFrame.EventProcPrefix
ms.assetid: 20d82dc1-6bb4-0338-6bfb-ce801825634d
ms.date: 06/08/2017
---


# BoundObjectFrame.EventProcPrefix Property (Access)

Gets or sets the prefix portion of an event procedure name. Read/write  **String**.


## Syntax

 _expression_. **EventProcPrefix**

 _expression_ A variable that represents a **BoundObjectFrame** object.


## Remarks

For example, if you have a command button with an event procedure named Details_Click, the  **EventProcPrefix** property returns the string "Details".

Microsoft Access adds the prefix portion of an event procedure name to the event name with an underscore character (_).


## See also


#### Concepts


[BoundObjectFrame Object](boundobjectframe-object-access.md)

