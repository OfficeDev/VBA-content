---
title: NavigationButton.OnPush Property (Access)
keywords: vbaac10.chm10459
f1_keywords:
- vbaac10.chm10459
ms.prod: access
api_name:
- Access.NavigationButton.OnPush
ms.assetid: 09b91008-ab58-d7d6-40af-aeb2e41d8823
ms.date: 06/08/2017
---


# NavigationButton.OnPush Property (Access)

Sets or returns the value of the  **On Click** box in the **Properties** window of a command button. Read/write **String**.


## Syntax

 _expression_. **OnPush**

 _expression_ A variable that represents a **NavigationButton** object.


## Remarks

This property is helpful for programmatically changing the action Microsoft Access takes when an event is triggered. For example, between event calls you may want to change an expression's parameters, or switch from an event procedure to an expression or macro, depending on the circumstances under which the event was triggered. 

Although the name of this property is  **OnPush**, setting this property actually sets the value of the **On Click** box.

The  **OnPush** value will be one of the following, depending on the selection chosen in the **Choose Builder** window (accessed by clicking the **Build** button next to the **On Click** box in the command button's **Properties** window):


- If Expression Builder is chosen, the value will be "= _expression_ ", where _expression_ is the expression from the Expression Builder window.
    
- If Macro Builder is chosen, the value is the name of the macro. 
    
- If Code Builder is chosen, the value will be "[Event Procedure]". 
    
If the  **On Click** box is blank, the property value is an empty string.


## Example

The following example prints the value of the  **OnPush** property in the Immediate window for the "OK" button on the "Order Entry" form.


```vb
Debug.Print Forms("Order Entry").Controls("OK").OnPush
```


## See also


#### Concepts


[NavigationButton Object](navigationbutton-object-access.md)

