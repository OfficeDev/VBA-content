---
title: Section.OnMouseMove Property (Access)
keywords: vbaac10.chm12210
f1_keywords:
- vbaac10.chm12210
ms.prod: access
api_name:
- Access.Section.OnMouseMove
ms.assetid: f68e19c8-1eeb-7edc-0296-c4eadb313125
ms.date: 06/08/2017
---


# Section.OnMouseMove Property (Access)

Sets or returns the value of the  **On Mouse Move** box in the **Properties** window. Read/write **String**.


## Syntax

 _expression_. **OnMouseMove**

 _expression_ A variable that represents a **Section** object.


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


[Section Object](section-object-access.md)

