---
title: Form.OnClose Property (Access)
keywords: vbaac10.chm13444
f1_keywords:
- vbaac10.chm13444
ms.prod: access
api_name:
- Access.Form.OnClose
ms.assetid: af4a7532-f12a-5194-9636-a09f9221f465
ms.date: 06/08/2017
---


# Form.OnClose Property (Access)

Sets or returns the value of the  **On Close** box in the **Properties** window of a form or report. Read/write **String**.


## Syntax

 _expression_. **OnClose**

 _expression_ A variable that represents a **Form** object.


## Remarks

This property is helpful for programmatically changing the action Microsoft Access takes when an event is triggered. For example, between event calls you may want to change an expression's parameters, or switch from an event procedure to an expression or macro, depending on the circumstances under which the event was triggered.

The  **Close** event occurs when when a form or report is closed and removed from the screen.

The  **OnClose** value will be one of the following, depending on the selection chosen in the **Choose Builder** window (accessed by clicking the **Build** button next to the **On Close** box in the object's **Properties** window):


- If Expression Builder is chosen, the value will be "= _expression_", where  _expression_ is the expression from the Expression Builder window.
    
- If Macro Builder is chosen, the value is the name of the macro. 
    
- If Code Builder is chosen, the value will be "[Event Procedure]". 
    
If the  **On Close** box is blank, the property value is an empty string.


## See also


#### Concepts


[Form Object](form-object-access.md)

