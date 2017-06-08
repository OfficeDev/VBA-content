---
title: CommandButton.Caption Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 80dfbd73-6315-cb39-ad64-e87cf9083277
ms.date: 06/08/2017
---


# CommandButton.Caption Property (Outlook Forms Script)

Returns or sets a  **String** that appears on the button to describe what it does. Read/write.


## Syntax

 _expression_. **Caption**

 _expression_A variable that represents a  **CommandButton** object.


## Remarks

The default caption for a control is a unique name based on the type of control. For example, CommandButton1 is the default caption for the first command button in a form.

If a control's caption is too long, the caption is truncated. If a form's caption is too long for the title bar, the title is displayed with an ellipsis.

The  **[ForeColor](commandbutton-forecolor-property-outlook-forms-script.md)** property of the control determines the color of the text in the caption.

Setting  **[AutoSize](commandbutton-autosize-property-outlook-forms-script.md)** to **True** automatically adjusts the size of the control to frame the entire caption.


