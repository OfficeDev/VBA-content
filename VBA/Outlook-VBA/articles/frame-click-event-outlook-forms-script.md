---
title: Frame.Click Event (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 1ca51b8c-2756-b73a-bebe-c4093bfdbbce
ms.date: 06/08/2017
---


# Frame.Click Event (Outlook Forms Script)

Occurs when the user clicks inside the control.


## Syntax

 _expression_. **Click**

 _expression_A variable that represents a  **Frame** object.


## Remarks

The following are examples of actions that initiate the  **Click** event of the specified control:


- Clicking a blank area of a form or a disabled control (other than a list box) on the form.
    
- Clicking a control with the left mouse button (left-clicking).
    
- Pressing a control's accelerator key.
    


Left-clicking changes the value of a control, thus it initiates the  **Click** event. Right-clicking does not change the value of the control, so it does not initiate the **Click** event.


