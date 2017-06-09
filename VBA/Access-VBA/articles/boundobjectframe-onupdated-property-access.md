---
title: BoundObjectFrame.OnUpdated Property (Access)
keywords: vbaac10.chm10963
f1_keywords:
- vbaac10.chm10963
ms.prod: access
api_name:
- Access.BoundObjectFrame.OnUpdated
ms.assetid: 1af7adce-8d59-d8ac-cd3a-102266e55618
ms.date: 06/08/2017
---


# BoundObjectFrame.OnUpdated Property (Access)

Sets or returns the value of the  **On Updated** box in the **Properties** window of a form or report. Read/write **String**.


## Syntax

 _expression_. **OnUpdated**

 _expression_ A variable that represents a **BoundObjectFrame** object.


## Remarks

This property is helpful for programmatically changing the action Microsoft Access takes when an event is triggered. For example, between event calls you may want to change an expression's parameters, or switch from an event procedure to an expression or macro, depending on the circumstances under which the event was triggered.

The  **OnUpdated** value will be one of the following, depending on the selection chosen in the **Choose Builder** window (accessed by clicking the **Build** button next to the **On Updated** box in the object's **Properties** window):


- If Expression Builder is chosen, the value will be "= _expression_", where  _expression_ is the expression from the Expression Builder window.
    
- If Macro Builder is chosen, the value is the name of the macro. 
    
- If Code Builder is chosen, the value will be "[Event Procedure]". 
    
If the  **On Updated** box is blank, the property value is an empty string.


## See also


#### Concepts


[BoundObjectFrame Object](boundobjectframe-object-access.md)

