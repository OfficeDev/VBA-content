---
title: CommandButton.Click Event (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 747d6f8f-c4da-f670-d476-21729387c4bc
ms.date: 06/08/2017
---


# CommandButton.Click Event (Outlook Forms Script)

Occurs when the user clicks inside the control.


## Syntax

 _expression_. **Click**

 _expression_A variable that represents a  **CommandButton** object.


## Remarks

The following are examples of actions that initiate the  **Click** event of the specified control:


- Clicking a blank area of a form or a disabled control (other than a list box) on the form.
    
- Clicking a  **[CommandButton](commandbutton-object-outlook-forms-script.md)**.
    
- Pressing the  **SPACEBAR** when a **CommandButton** has the focus.
    
- Clicking a control with the left mouse button (left-clicking).
    
- Pressing  **ENTER** on a form that has a command button whose **Default** property is set to **True**, as long as no other command button has the focus.
    
- Pressing  **ESC** on a form that has a command button whose **Cancel** property is set to **True**, as long as no other command button has the focus.
    
- Pressing a control's accelerator key.
    


Left-clicking changes the value of a control, thus it initiates the  **Click** event. Right-clicking does not change the value of the control, so it does not initiate the **Click** event.


