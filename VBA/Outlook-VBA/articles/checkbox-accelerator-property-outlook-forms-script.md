---
title: CheckBox.Accelerator Property (Outlook Forms Script)
keywords: olfm10.chm2000690
f1_keywords:
- olfm10.chm2000690
ms.prod: outlook
ms.assetid: 940cec9e-8c29-4db9-77bd-b52cee7748f9
ms.date: 06/08/2017
---


# CheckBox.Accelerator Property (Outlook Forms Script)

Returns or sets the accelerator key for a control. Read/write.


## Syntax

 _expression_. **Accelerator**

 _expression_A variable that represents a  **CheckBox** object.


## Remarks

To designate an accelerator key, enter a single character for the  **Accelerator** property. You can set **Accelerator** in the control's property sheet or in code. If the value of this property contains more than one character, the first character in the string becomes the value of **Accelerator**. You cannot use digits in an accelerator.

When an accelerator key is used, there is no visual feedback (other than focus) to indicate that the control initiated the  **[Click](checkbox-click-event-outlook-forms-script.md)** event. For example, if the accelerator key applies to a **[CommandButton](commandbutton-object-outlook-forms-script.md)**, the user will not see the button pressed in the interface. The button receives the focus, however, when the user presses the accelerator key.


