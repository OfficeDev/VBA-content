---
title: Label.OnMouseMove Property (Access)
keywords: vbaac10.chm10227
f1_keywords:
- vbaac10.chm10227
ms.prod: access
api_name:
- Access.Label.OnMouseMove
ms.assetid: 31d14519-05c4-687f-af51-8ca198f5c7de
ms.date: 06/08/2017
---


# Label.OnMouseMove Property (Access)

Sets or returns the value of the  **On Mouse Move** box in the **Properties** window. Read/write **String**.


## Syntax

 _expression_. **OnMouseMove**

 _expression_ A variable that represents a **Label** object.


## Remarks

This property is helpful for programmatically changing the action Microsoft Access takes when an event is triggered. For example, between event calls you may want to change an expression's parameters, or switch from an event procedure to an expression or macro, depending on the circumstances under which the event was triggered. 

The  **MouseMove** event occurs when the user moves the mouse over the object.

The  **OnMouseMove** value will be one of the following, depending on the selection chosen in the **Choose Builder** window (accessed by clicking the **Build** button next to the **On Mouse Move** box in the object's **Properties** window):


- If Expression Builder is chosen, the value will be "= _expression_", where  _expression_ is the expression from the Expression Builder window.
    
- If Macro Builder is chosen, the value is the name of the macro. 
    
- If Code Builder is chosen, the value will be "[Event Procedure]". 
    
If the  **On Mouse Move** box is blank, the property value is an empty string.


## Example

The following example prints the value of the  **OnMouseMove** property in the Immediate window for the button named "OK" on the "Order Entry" form.


```vb
Debug.Print Forms("Order Entry").Controls("OK").OnMouseMove
```


## See also


#### Concepts


[Label Object](label-object-access.md)

