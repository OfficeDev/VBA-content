---
title: Form.OnDeactivate Property (Access)
keywords: vbaac10.chm13446
f1_keywords:
- vbaac10.chm13446
ms.prod: access
api_name:
- Access.Form.OnDeactivate
ms.assetid: c241c3cc-377b-7407-87f3-3003edb3ff8f
ms.date: 06/08/2017
---


# Form.OnDeactivate Property (Access)

Sets or returns the value of the  **On Deactivate** box in the **Properties** window of a form or report. Read/write **String**.


## Syntax

 _expression_. **OnDeactivate**

 _expression_ A variable that represents a **Form** object.


## Remarks

This property is helpful for programmatically changing the action Microsoft Access takes when an event is triggered. For example, between event calls you may want to change an expression's parameters, or switch from an event procedure to an expression or macro, depending on the circumstances under which the event was triggered.

The  **Deactivate** event occurs when the form or report loses the focus to a Table, Query, Form, Report, Macro, or Module window, or to the Database window.

The  **OnDeactivate** value will be one of the following, depending on the selection chosen in the **Choose Builder** window (accessed by clicking the **Build** button next to the **On Deactivate** box in the object's **Properties** window):


- If Expression Builder is chosen, the value will be "= _expression_", where  _expression_ is the expression from the Expression Builder window.
    
- If Macro Builder is chosen, the value is the name of the macro. 
    
- If Code Builder is chosen, the value will be "[Event Procedure]". 
    
If the  **On Deactivate** box is blank, the property value is an empty string.


## Example

The following example associates the  **Deactivate** event with the macro "Deactivate_Macro" for the "Order Entry" form.


```vb
Forms("Order Entry").OnDeactivate = "Deactivate_Macro"
```


## See also


#### Concepts


[Form Object](form-object-access.md)

