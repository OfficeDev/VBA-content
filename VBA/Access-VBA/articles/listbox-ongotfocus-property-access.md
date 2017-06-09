---
title: ListBox.OnGotFocus Property (Access)
keywords: vbaac10.chm11282
f1_keywords:
- vbaac10.chm11282
ms.prod: access
api_name:
- Access.ListBox.OnGotFocus
ms.assetid: ded32fef-1821-5288-9265-c1e1f4819efe
ms.date: 06/08/2017
---


# ListBox.OnGotFocus Property (Access)

Sets or returns the value of the  **On Got Focus** box in the **Properties** window of the specified object. Read/write **String**.


## Syntax

 _expression_. **OnGotFocus**

 _expression_ A variable that represents a **ListBox** object.


## Remarks

This property is helpful for programmatically changing the action Microsoft Access takes when an event is triggered. For example, between event calls you may want to change an expression's parameters, or switch from an event procedure to an expression or macro, depending on the circumstances under which the event was triggered. 

The  **GotFocus** event occurs when the object receives the focus.

The  **OnGotFocus** value will be one of the following, depending on the selection chosen in the **Choose Builder** window (accessed by clicking the **Build** button next to the **On Got Focus** box in the object's **Properties** window):


- If Expression Builder is chosen, the value will be "= _expression_ ", where _expression_ is the expression from the Expression Builder window.
    
- If Macro Builder is chosen, the value is the name of the macro. 
    
- If Code Builder is chosen, the value will be "[Event Procedure]". 
    
If the  **On Got Focus** box is blank, the property value is an empty string.


## Example

The following example prints the value of the  **OnGotFocus** property in the Immediate window for the button named "OK" on the "Order Entry" form.


```vb
Debug.Print Forms("Order Entry").Controls("OK").OnGotFocus
```


## See also


#### Concepts


[ListBox Object](listbox-object-access.md)

