---
title: Form.OnDelete Property (Access)
keywords: vbaac10.chm13437,vbaac10.chm5427
f1_keywords:
- vbaac10.chm13437,vbaac10.chm5427
ms.prod: access
api_name:
- Access.Form.OnDelete
ms.assetid: 97cfb9eb-e1c7-a879-a8aa-d26ff337efbb
ms.date: 06/08/2017
---


# Form.OnDelete Property (Access)

Sets or returns the value of the  **On Delete** box in the **Properties** window of a form. Read/write **String**.


## Syntax

 _expression_. **OnDelete**

 _expression_ A variable that represents a **Form** object.


## Remarks

This property is helpful for programmatically changing the action Microsoft Access takes when an event is triggered. For example, between event calls you may want to change an expression's parameters, or switch from an event procedure to an expression or macro, depending on the circumstances under which the event was triggered. 

The  **Delete** event occurs when the user performs some action, such as pressing the DELETE key to delete a record, but before the record is actually deleted.

The  **OnDelete** value will be one of the following, depending on the selection chosen in the **Choose Builder** window (accessed by clicking the **Build** button next to the **On Delete** box in the form's **Properties** window):


- If Expression Builder is chosen, the value will be "= _expression_ ", where _expression_ is the expression from the Expression Builder window.
    
- If Macro Builder is chosen, the value is the name of the macro. 
    
- If Code Builder is chosen, the value will be "[Event Procedure]". 
    
If the  **On Delete** box is blank, the property value is an empty string.


## Example

The following example associates the  **Delete** event with the "Form_Delete" event for the "Order Entry" form.


```vb
Forms("Order Entry").OnDelete = "[Event Procedure]"
```


## See also


#### Concepts


[Form Object](form-object-access.md)

