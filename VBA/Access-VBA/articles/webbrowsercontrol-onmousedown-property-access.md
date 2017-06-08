---
title: WebBrowserControl.OnMouseDown Property (Access)
keywords: vbaac10.chm14380
f1_keywords:
- vbaac10.chm14380
ms.prod: access
api_name:
- Access.WebBrowserControl.OnMouseDown
ms.assetid: b3a1d803-078c-fdd9-7e1f-4ce93bd1719c
ms.date: 06/08/2017
---


# WebBrowserControl.OnMouseDown Property (Access)

Sets or returns the value of the  **On Mouse Down** box in the **Properties** window. Read/write **String**.


## Syntax

 _expression_. **OnMouseDown**

 _expression_ A variable that represents a **WebBrowserControl** object.


## Remarks

This property is helpful for programmatically changing the action Microsoft Access takes when an event is triggered. For example, between event calls you may want to change an expression's parameters, or switch from an event procedure to an expression or macro, depending on the circumstances under which the event was triggered. 

The  **MouseDown** event occurs when the user clicks the mouse button while the mouse pointer rests over the object.

The  **OnMouseDown** value will be one of the following, depending on the selection chosen in the **Choose Builder** window (accessed by clicking the **Build** button next to the **On Mouse Down** box in the object's **Properties** window):


- If Expression Builder is chosen, the value will be "= _expression_", where  _expression_ is the expression from the Expression Builder window.
    
- If Macro Builder is chosen, the value is the name of the macro. 
    
- If Code Builder is chosen, the value will be "[Event Procedure]". 
    
If the  **On Mouse Down** box is blank, the property value is an empty string.


## Example

The following example prints the value of the  **OnMouseDown** property in the Immediate window for the button named "OK" on the "Order Entry" form.


```vb
Debug.Print Forms("Order Entry").Controls("OK").OnMouseDown
```


## See also


#### Concepts


[WebBrowserControl Object](webbrowsercontrol-object-access.md)

