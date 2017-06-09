---
title: ComboBox.OnNotInList Property (Access)
keywords: vbaac10.chm11450,vbaac10.chm4100
f1_keywords:
- vbaac10.chm11450,vbaac10.chm4100
ms.prod: access
api_name:
- Access.ComboBox.OnNotInList
ms.assetid: 307e9f0c-6db7-b995-166b-060c697b9f6e
ms.date: 06/08/2017
---


# ComboBox.OnNotInList Property (Access)

Sets or returns the value of the  **On Not in List** box in the **Properties** window of a combo box. Read/write **String**.


## Syntax

 _expression_. **OnNotInList**

 _expression_ A variable that represents a **ComboBox** object.


## Remarks

This property is helpful for programmatically changing the action Microsoft Access takes when an event is triggered. For example, between event calls you may want to change an expression's parameters, or switch from an event procedure to an expression or macro, depending on the circumstances under which the event was triggered. 

The  **NotInList** event occurs when the user enters a value in the text box portion of a combo box that isn't in the combo box list.

The  **OnNotInList** value will be one of the following, depending on the selection chosen in the **Choose Builder** window (accessed by clicking the **Build** button next to the **On Not in List** box in the combo box's **Properties** window):


- If Expression Builder is chosen, the value will be "= _expression_ ", where _expression_ is the expression from the Expression Builder window.
    
- If Macro Builder is chosen, the value is the name of the macro. 
    
- If Code Builder is chosen, the value will be "[Event Procedure]". 
    
If the  **On Not in List** box is blank, the property value is an empty string.


## Example

The following example prints the value of the  **OnNotInList** property in the Immediate window for the "State" combo box in the "Order Entry" form.


```vb
Debug.Print Forms("Order Entry").Controls("State").OnNotInList
```


## See also


#### Concepts


[ComboBox Object](combobox-object-access.md)

