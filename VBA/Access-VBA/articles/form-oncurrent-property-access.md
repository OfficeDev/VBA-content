---
title: Form.OnCurrent Property (Access)
keywords: vbaac10.chm13430
f1_keywords:
- vbaac10.chm13430
ms.prod: access
api_name:
- Access.Form.OnCurrent
ms.assetid: bb7eb7be-7bb6-8fdd-6a48-f5b33ad7dc14
ms.date: 06/08/2017
---


# Form.OnCurrent Property (Access)

Sets or returns the value of the  **On Current** box in the **Properties** window of a form. Read/write **String**.


## Syntax

 _expression_. **OnCurrent**

 _expression_ A variable that represents a **Form** object.


## Remarks

This property is helpful for programmatically changing the action Microsoft Access takes when an event is triggered. For example, between event calls you may want to change an expression's parameters, or switch from an event procedure to an expression or macro, depending on the circumstances under which the event was triggered. 

The  **Current** event occurs when the focus moves to a record, making it the current record, or when the form is refreshed or requeried.

The  **OnCurrent** value will be one of the following, depending on the selection chosen in the **Choose Builder** window (accessed by clicking the **Build** button next to the **On Apply Filter** box in the form's **Properties** window):


- If Expression Builder is chosen, the value will be "= _expression_ ", where _expression_ is the expression from the Expression Builder window.
    
- If Macro Builder is chosen, the value is the name of the macro. 
    
- If Code Builder is chosen, the value will be "[Event Procedure]". 
    
If the  **On Current** box is blank, the property value is an empty string.


## Example

The following example associates the  **Current** event with the macro "Current_Macro" for the "Order Entry" form.


```vb
Forms("Order Entry").OnDeactivate = "Current_Macro" 

```


## See also


#### Concepts


[Form Object](form-object-access.md)

