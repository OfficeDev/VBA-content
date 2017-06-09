---
title: ObjectFrame.OnEnter Property (Access)
keywords: vbaac10.chm11615
f1_keywords:
- vbaac10.chm11615
ms.prod: access
api_name:
- Access.ObjectFrame.OnEnter
ms.assetid: 386524c3-8208-05dd-5d0f-9899e4619eb7
ms.date: 06/08/2017
---


# ObjectFrame.OnEnter Property (Access)

Sets or returns the value of the  **On Enter** box in the **Properties** window of specified object. Read/write **String**. .


## Syntax

 _expression_. **OnEnter**

 _expression_ A variable that represents an **ObjectFrame** object.


## Remarks

This property is helpful for programmatically changing the action Microsoft Access takes when an event is triggered. For example, between event calls you may want to change an expression's parameters, or switch from an event procedure to an expression or macro, depending on the circumstances under which the event was triggered. 

The  **Enter** event occurs before a control actually receives the focus from a control on the same form.

The  **OnEnter** value will be one of the following, depending on the selection chosen in the **Choose Builder** window (accessed by clicking the **Build** button next to the **On Enter** box in the object's **Properties** window):


- If Expression Builder is chosen, the value will be "= _expression_ ", where _expression_ is the expression from the Expression Builder window.
    
- If Macro Builder is chosen, the value is the name of the macro. 
    
- If Code Builder is chosen, the value will be "[Event Procedure]". 
    
If the  **On Enter** box is blank, the property value is an empty string.


## Example

The following example associates the  **Enter** event with the macro "Enter_Macro" for the button named "OK" on the "Order Entry" form.


```vb
Forms("Order Entry").Controls("OK").OnEnter = "Enter_Macro"
```


## See also


#### Concepts


[ObjectFrame Object](objectframe-object-access.md)

