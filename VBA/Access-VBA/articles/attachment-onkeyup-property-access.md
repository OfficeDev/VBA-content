---
title: Attachment.OnKeyUp Property (Access)
keywords: vbaac10.chm13951
f1_keywords:
- vbaac10.chm13951
ms.prod: access
api_name:
- Access.Attachment.OnKeyUp
ms.assetid: 56e5a246-5907-f537-0c89-a746beab0865
ms.date: 06/08/2017
---


# Attachment.OnKeyUp Property (Access)

Sets or returns the value of the  **On Key Up** box in the **Properties** window. Read/write **String**.


## Syntax

 _expression_. **OnKeyUp**

 _expression_ A variable that represents an **Attachment** object.


## Remarks

This property is helpful for programmatically changing the action Microsoft Access takes when an event is triggered. For example, between event calls you may want to change an expression's parameters, or switch from an event procedure to an expression or macro, depending on the circumstances under which the event was triggered. 

The  **KeyUp** event occurs when a user presses a key while a form or control has the focus. This event also occurs if you send a keystroke to a form or control by using the SendKeys action in a macro or the **SendKeys** statement in Visual Basic.

The  **OnKeyUp** value will be one of the following, depending on the selection chosen in the **Choose Builder** window (accessed by clicking the **Build** button next to the **On Key Up** box in the object's **Properties** window):


- If Expression Builder is chosen, the value will be "= _expression_ ", where _expression_ is the expression from the Expression Builder window.
    
- If Macro Builder is chosen, the value is the name of the macro. 
    
- If Code Builder is chosen, the value will be "[Event Procedure]". 
    
If the  **On Key Up** box is blank, the property value is an empty string.


## See also


#### Concepts


[Attachment Object](attachment-object-access.md)

