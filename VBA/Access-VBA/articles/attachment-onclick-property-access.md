---
title: Attachment.OnClick Property (Access)
keywords: vbaac10.chm13945
f1_keywords:
- vbaac10.chm13945
ms.prod: access
api_name:
- Access.Attachment.OnClick
ms.assetid: c1d1ddcb-db80-e0e1-4318-0cf9477d7316
ms.date: 06/08/2017
---


# Attachment.OnClick Property (Access)

Sets or returns the value of the  **On Click** box in the **Properties** window. Read/write **String**.


## Syntax

 _expression_. **OnClick**

 _expression_ A variable that represents an **Attachment** object.


## Remarks

This property is helpful for programmatically changing the action Microsoft Access takes when an event is triggered. For example, between event calls you may want to change an expression's parameters, or switch from an event procedure to an expression or macro, depending on the circumstances under which the event was triggered. 

The  **Click** event occurs when a user presses and releases the left mouse button over an object.

The  **OnClick** value will be one of the following, depending on the selection chosen in the **Choose Builder** window (accessed by clicking the **Build** button next to the **On Click** box in the object's **Properties** window):


- If Expression Builder is chosen, the value will be "= _expression_", where  _expression_ is the expression from the Expression Builder window.
    
- If Macro Builder is chosen, the value is the name of the macro. 
    
- If Code Builder is chosen, the value will be "[Event Procedure]". 
    
If the  **On Click** box is blank, the property value is an empty string.


## See also


#### Concepts


[Attachment Object](attachment-object-access.md)

