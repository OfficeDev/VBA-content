---
title: ComboBox.OnClick Property (Access)
keywords: vbaac10.chm11455
f1_keywords:
- vbaac10.chm11455
ms.prod: access
api_name:
- Access.ComboBox.OnClick
ms.assetid: 9cb266d8-6e7d-80c9-c5e9-1d2406b7d54d
ms.date: 06/08/2017
---


# ComboBox.OnClick Property (Access)

Sets or returns the value of the  **On Click** box in the **Properties** window. Read/write **String**.


## Syntax

 _expression_. **OnClick**

 _expression_ A variable that represents a **ComboBox** object.


## Remarks

This property is helpful for programmatically changing the action Microsoft Access takes when an event is triggered. For example, between event calls you may want to change an expression's parameters, or switch from an event procedure to an expression or macro, depending on the circumstances under which the event was triggered. 

The  **Click** event occurs when one of the following things occurs:


- The user presses and releases the left mouse button over combo box.
    
-  The user commits a change to the value of the combo box. In this case, the **Click** event fires after the **AfterUpdate** event and before the **Exit** event.
    
The  **OnClick** value will be one of the following, depending on the selection chosen in the **Choose Builder** window (accessed by clicking the **Build** button next to the **On Click** box in the object's **Properties** window):


- If Expression Builder is chosen, the value will be "= _expression_", where  _expression_ is the expression from the Expression Builder window.
    
- If Macro Builder is chosen, the value is the name of the macro. 
    
- If Code Builder is chosen, the value will be "[Event Procedure]". 
    
If the  **On Click** box is blank, the property value is an empty string.


## Example

The following example associates the  **Click** event with the "OK_Click" event procedure for the button named "OK" on the "Order Entry" form, if there is currently no association.


```vb
With Forms("Order Entry").Controls("OK") 
 If .OnClick = "" Then 
 .OnClick = "[Event Procedure]" 
 End If 
End With 

```


## See also


#### Concepts


[ComboBox Object](combobox-object-access.md)

