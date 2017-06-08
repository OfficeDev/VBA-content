---
title: CustomControl.OnUpdated Property (Access)
keywords: vbaac10.chm12040
f1_keywords:
- vbaac10.chm12040
ms.prod: access
api_name:
- Access.CustomControl.OnUpdated
ms.assetid: 6cd30c42-d645-6ca8-5c9e-7a5951283fd9
ms.date: 06/08/2017
---


# CustomControl.OnUpdated Property (Access)

Sets or returns the value of the  **On Updated** box in the **Properties** window of a form or report. Read/write **String**.


## Syntax

 _expression_. **OnUpdated**

 _expression_ A variable that represents a **CustomControl** object.


## Remarks

This property is helpful for programmatically changing the action Microsoft Access takes when an event is triggered. For example, between event calls you may want to change an expression's parameters, or switch from an event procedure to an expression or macro, depending on the circumstances under which the event was triggered.

The  **OnUpdated** value will be one of the following, depending on the selection chosen in the **Choose Builder** window (accessed by clicking the **Build** button next to the **On Updated** box in the object's **Properties** window):


- If Expression Builder is chosen, the value will be "= _expression_", where  _expression_ is the expression from the Expression Builder window.
    
- If Macro Builder is chosen, the value is the name of the macro. 
    
- If Code Builder is chosen, the value will be "[Event Procedure]". 
    
If the  **On Updated** box is blank, the property value is an empty string.


## See also


#### Concepts


[CustomControl Object](customcontrol-object-access.md)

