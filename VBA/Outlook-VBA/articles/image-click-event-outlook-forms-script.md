---
title: Image.Click Event (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 59ac08ce-2527-6cfb-ac0b-66322bc10e9f
ms.date: 06/08/2017
---


# Image.Click Event (Outlook Forms Script)

Occurs when the user clicks inside the control.


## Syntax

 _expression_. **Click**

 _expression_A variable that represents an  **Image** object.


## Remarks

The following are examples of actions that initiate the  **Click** event of the specified control:


- Clicking a blank area of a form or a disabled control (other than a list box) on the form.
    
- Clicking a control with the left mouse button (left-clicking).
    
- Pressing a control's accelerator key.
    


Left-clicking changes the value of a control, thus it initiates the  **Click** event. Right-clicking does not change the value of the control, so it does not initiate the **Click** event.


