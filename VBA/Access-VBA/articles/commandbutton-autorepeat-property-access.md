---
title: CommandButton.AutoRepeat Property (Access)
keywords: vbaac10.chm10457,vbaac10.chm4276
f1_keywords:
- vbaac10.chm10457,vbaac10.chm4276
ms.prod: access
api_name:
- Access.CommandButton.AutoRepeat
ms.assetid: 028a5bdd-1e37-0499-202f-c9e3fdb83838
ms.date: 06/08/2017
---


# CommandButton.AutoRepeat Property (Access)

You can use the  **AutoRepeat** property to specify whether an event procedure or macro runs repeatedly while a command button on a form remains pressed in. Read/write **Boolean**.


## Syntax

 _expression_. **AutoRepeat**

 _expression_ A variable that represents a **CommandButton** object.


## Remarks

The default value is  **False**.

The initial repeat of the event procedure or macro occurs 0.5 seconds after it first runs. Subsequent repeats occur either 0.25 seconds apart or the duration of the event procedure or macro, whichever is longer.

If the code attached to the command button causes the current record to change, the  **AutoRepeat** property has no effect.

If the code attached to the command button causes changes to another control on a form, use the  **DoEvents** function to ensure proper screen updating.


## See also


#### Concepts


[CommandButton Object](commandbutton-object-access.md)

