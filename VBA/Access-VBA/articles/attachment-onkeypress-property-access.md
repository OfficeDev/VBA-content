---
title: Attachment.OnKeyPress Property (Access)
keywords: vbaac10.chm13952
f1_keywords:
- vbaac10.chm13952
ms.prod: access
api_name:
- Access.Attachment.OnKeyPress
ms.assetid: 9f3213eb-9c37-f550-6c14-e6dd85d030a5
ms.date: 06/08/2017
---


# Attachment.OnKeyPress Property (Access)

Sets or returns the value of the  **On Key Press** box in the **Properties** window. Read/write **String**.


## Syntax

 _expression_. **OnKeyPress**

 _expression_ A variable that represents an **Attachment** object.


## Remarks

This property is helpful for programmatically changing the action Microsoft Access takes when an event is triggered. For example, between event calls you may want to change an expression's parameters, or switch from an event procedure to an expression or macro, depending on the circumstances under which the event was triggered. 

The  **KeyPress** event occurs when a user presses a key while a form or control has the focus. This event also occurs if you send a keystroke to a form or control by using the SendKeys action in a macro or the **SendKeys** statement in Visual Basic.

The  **OnKeyPress** value will be one of the following, depending on the selection chosen in the **Choose Builder** window (accessed by clicking the **Build** button next to the **On Key Press** box in the object's **Properties** window):


- If Expression Builder is chosen, the value will be "= _expression_ ", where _expression_ is the expression from the Expression Builder window.
    
- If Macro Builder is chosen, the value is the name of the macro. 
    
- If Code Builder is chosen, the value will be "[Event Procedure]". 
    
If the  **On Key Press** box is blank, the property value is an empty string.


## See also


#### Concepts


[Attachment Object](attachment-object-access.md)

