---
title: Frame.Caption Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 6075400e-e4c0-1a1c-dea1-8628d191337b
ms.date: 06/08/2017
---


# Frame.Caption Property (Outlook Forms Script)

Returns or sets a  **String** that appears on an object to identify or describe it. Read/write.


## Syntax

 _expression_. **Caption**

 _expression_A variable that represents a  **Frame** object.


## Remarks

The default caption for a control is a unique name based on the type of control. For example, CommandButton1 is the default caption for the first command button in a form.

If a control's caption is too long, the caption is truncated. If a form's caption is too long for the title bar, the title is displayed with an ellipsis.

The  **[ForeColor](frame-forecolor-property-outlook-forms-script.md)** property of the control determines the color of the text in the caption.


