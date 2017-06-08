---
title: ToggleButton.Click Event (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 777a0efc-c376-221b-ecea-5bd7797488de
ms.date: 06/08/2017
---


# ToggleButton.Click Event (Outlook Forms Script)

Occurs when the user definitively selects a value for the control that has more than one possible value.


## Syntax

 _expression_. **Click**

 _expression_A variable that represents a  **ToggleButton** object.


## Remarks

The following are examples of actions that initiate the  **Click** event of the specified control:


- Clicking a blank area of a form or a disabled control (other than a list box) on the form.
    
- Clicking a control with the left mouse button (left-clicking).
    
- Pressing a control's accelerator key.
    


For some controls, the  **Click** event occurs when the **Value** property changes. However, using the **PropertyChange** or **CustomPropertyChange** event is the preferred technique for detecting a new value for a property. The following are examples of actions that initiate the **Click** event due to assigning a new value to a control: clicking a **[ToggleButton](togglebutton-object-outlook-forms-script.md)**, pressing the SPACEBAR when a toggle button has the focus, pressing the accelerator key, or changing the value of the control in code.

The  **Click** event is not initiated when **Value** is set to **Null**.

Left-clicking changes the value of a control, thus it initiates the  **Click** event. Right-clicking does not change the value of the control, so it does not initiate the **Click** event.


