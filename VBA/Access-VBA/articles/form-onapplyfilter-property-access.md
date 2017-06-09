---
title: Form.OnApplyFilter Property (Access)
keywords: vbaac10.chm13460
f1_keywords:
- vbaac10.chm13460
ms.prod: access
api_name:
- Access.Form.OnApplyFilter
ms.assetid: 5e147a50-5516-f6d3-c1c9-e2c4522cb804
ms.date: 06/08/2017
---


# Form.OnApplyFilter Property (Access)

Sets or returns the value of the  **On Apply Filter** box in the **Properties** window of a form. Read/write **String**.


## Syntax

 _expression_. **OnApplyFilter**

 _expression_ A variable that represents a **Form** object.


## Remarks

This property is helpful for programmatically changing the action Microsoft Access takes when an event is triggered. For example, between event calls you may want to change an expression's parameters, or switch from an event procedure to an expression or macro, depending on the circumstances under which the event was triggered. 

The  **Apply Filter** event occurs when a filter is applied or removed.

The  **OnApplyFilter** value will be one of the following, depending on the selection chosen in the **Choose Builder** window (accessed by clicking the **Build** button next to the **On Apply Filter** box in the form's **Properties** window):


- If Expression Builder is chosen, the value will be "= _expression_ ", where _expression_ is the expression from the Expression Builder window.
    
- If Macro Builder is chosen, the value is the name of the macro. 
    
- If Code Builder is chosen, the value will be "[Event Procedure]". 
    
If the  **On Apply Filter** box is blank, the property value is an empty string.


## Example

The following example associates the  **OnApplyFilter** property for the "Order Entry" form to the event "Form_ApplyFilter".


```vb
Forms("Order Entry").OnFilter = "[Event Procedure]"
```


## See also


#### Concepts


[Form Object](form-object-access.md)

