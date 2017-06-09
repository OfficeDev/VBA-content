---
title: CustomControl.EventProcPrefix Property (Access)
keywords: vbaac10.chm12006
f1_keywords:
- vbaac10.chm12006
ms.prod: access
api_name:
- Access.CustomControl.EventProcPrefix
ms.assetid: 578dc1f6-0977-e8b9-e96f-ae3408118456
ms.date: 06/08/2017
---


# CustomControl.EventProcPrefix Property (Access)

Gets or sets the prefix portion of an event procedure name. Read/write  **String**.


## Syntax

 _expression_. **EventProcPrefix**

 _expression_ A variable that represents a **CustomControl** object.


## Remarks

For example, if you have a command button with an event procedure named Details_Click, the  **EventProcPrefix** property returns the string "Details".

Microsoft Access adds the prefix portion of an event procedure name to the event name with an underscore character (_).


## See also


#### Concepts


[CustomControl Object](customcontrol-object-access.md)

