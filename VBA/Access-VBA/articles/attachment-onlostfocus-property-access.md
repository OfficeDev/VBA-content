---
title: Attachment.OnLostFocus Property (Access)
keywords: vbaac10.chm13944
f1_keywords:
- vbaac10.chm13944
ms.prod: access
api_name:
- Access.Attachment.OnLostFocus
ms.assetid: 546d0491-ddb8-87d4-9f97-d68cfd96070c
ms.date: 06/08/2017
---


# Attachment.OnLostFocus Property (Access)

Sets or returns the value of the  **On Lost Focus** box in the **Properties** window of the specified object. Read/write **String**.


## Syntax

 _expression_. **OnLostFocus**

 _expression_ A variable that represents an **Attachment** object.


## Remarks

This property is helpful for programmatically changing the action Microsoft Access takes when an event is triggered. For example, between event calls you may want to change an expression's parameters, or switch from an event procedure to an expression or macro, depending on the circumstances under which the event was triggered. 

The  **LostFocus** event occurs when the object loses the focus.

The  **OnLostFocus** value will be one of the following, depending on the selection chosen in the **Choose Builder** window (accessed by clicking the **Build** button next to the **On Lost Focus** box in the object's **Properties** window):


- If Expression Builder is chosen, the value will be "= _expression_ ", where _expression_ is the expression from the Expression Builder window.
    
- If Macro Builder is chosen, the value is the name of the macro. 
    
- If Code Builder is chosen, the value will be "[Event Procedure]". 
    
If the  **On Lost Focus** box is blank, the property value is an empty string.


## See also


#### Concepts


[Attachment Object](attachment-object-access.md)

