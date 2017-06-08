---
title: CommandButton.Accelerator Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 896f426a-6bc7-c6cc-96a0-eb0106b027f1
ms.date: 06/08/2017
---


# CommandButton.Accelerator Property (Outlook Forms Script)

Returns or sets the accelerator key for a control. Read/write.


## Syntax

 _expression_. **Accelerator**

 _expression_A variable that represents a  **CommandButton** object.


## Remarks

To designate an accelerator key, enter a single character for the  **Accelerator** property. You can set **Accelerator** in the control's property sheet or in code. If the value of this property contains more than one character, the first character in the string becomes the value of **Accelerator**. You cannot use digits in an accelerator.

When an accelerator key is used, there is no visual feedback (other than focus) to indicate that the control initiated the  **[Click](commandbutton-click-event-outlook-forms-script.md)** event. For example, if the accelerator key applies to a **[CommandButton](commandbutton-object-outlook-forms-script.md)**, the user will not see the button pressed in the interface. The button receives the focus, however, when the user presses the accelerator key.


