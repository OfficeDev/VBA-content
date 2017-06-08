---
title: NavigationControl.Click Event (Access)
keywords: vbaac10.chm14201
f1_keywords:
- vbaac10.chm14201
ms.prod: access
api_name:
- Access.NavigationControl.Click
ms.assetid: c49b26bd-dbab-666a-ecc0-2b3137bb10a0
ms.date: 06/08/2017
---


# NavigationControl.Click Event (Access)

The  **Click** event occurs when the user presses and then releases a mouse button over an object.


## Syntax

 _expression_. **Click**

 _expression_ A variable that represents a **NavigationControl** object.


### Return Value

nothing


## Remarks


- This event applies to a control containing a hyperlink.
    
To run a macro or event procedure when this event occurs, set the  **OnClick** property to the name of the macro or to [Event Procedure].

For a control, this event occurs when the user:


- Clicks a control with the left mouse button. Clicking a control with the right or middle mouse button does not trigger this event.
    
- Clicks a control containing hyperlink data with the left mouse button. Clicking a control with the right or middle mouse button does not trigger this event. When the user moves the mouse pointer over a control containing hyperlink data, the mouse pointer changes to a "hand" icon. When the user clicks the mouse button, the hyperlink is activated, and then the  **Click** event occurs.
    
- Selects an item in a combo box or list box, either by pressing the arrow keys and then pressing the ENTER key or by clicking the mouse button.
    
- Presses SPACEBAR when a command button, check box, option button, or toggle button has the focus.
    
- Presses the ENTER key on a form that has a command button whose  **Default** property is set to Yes.
    
- Presses the ESC key on a form that has a command button whose  **Cancel** property is set to Yes.
    
- Presses a control's access key. For example, if a command button's  **Caption** property is set to &;Go, pressing ALT+G triggers the event.
    
Typically, you attach a  **Click** event procedure or macro to a command button to carry out commands and command-like actions. For the other applicable controls, use this event to trigger actions in response to one of the occurrences discussed earlier in this topic.

You can use a  **CancelEvent** action in a DblClick macro to cancel the second **Click** event. For more information, see the DblClick event topic.

To distinguish between the left, right, and middle mouse buttons, use the  **MouseDown** and **MouseUp** events.


## See also


#### Concepts


[NavigationControl Object](navigationcontrol-object-access.md)

