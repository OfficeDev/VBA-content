---
title: Form.OnOpen Property (Access)
keywords: vbaac10.chm13440
f1_keywords:
- vbaac10.chm13440
ms.prod: access
api_name:
- Access.Form.OnOpen
ms.assetid: 151b9103-a25d-a595-6cab-20b737909fa6
ms.date: 06/08/2017
---


# Form.OnOpen Property (Access)

Sets or returns the value of the  **On Open** box in the **Properties** window of a form or report. Read/write **String**.


## Syntax

 _expression_. **OnOpen**

 _expression_ A variable that represents a **Form** object.


## Remarks

This property is helpful for programmatically changing the action Microsoft Access takes when an event is triggered. For example, between event calls you may want to change an expression's parameters, or switch from an event procedure to an expression or macro, depending on the circumstances under which the event was triggered.

The  **Open** event occurs when a form is opened, but before the first record is displayed. For reports, the event occurs before a report is previewed or printed.

The  **OnOpen** value will be one of the following, depending on the selection chosen in the **Choose Builder** window (accessed by clicking the **Build** button next to the **On Open** box in the object's **Properties** window):


- If Expression Builder is chosen, the value will be "= _expression_", where  _expression_ is the expression from the Expression Builder window.
    
- If Macro Builder is chosen, the value is the name of the macro. 
    
- If Code Builder is chosen, the value will be "[Event Procedure]". 
    
If the  **On Open** box is blank, the property value is an empty string.


## See also


#### Concepts


[Form Object](form-object-access.md)

