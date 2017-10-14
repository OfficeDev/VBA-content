---
title: CheckBox.Caption Property (Outlook Forms Script)
keywords: olfm10.chm2000880
f1_keywords:
- olfm10.chm2000880
ms.prod: outlook
ms.assetid: ee143257-1e0d-d50a-7ed1-44a53af4a1c0
ms.date: 06/08/2017
---


# CheckBox.Caption Property (Outlook Forms Script)

Returns or sets a  **String** that appears on an object to identify or describe it. Read/write.


## Syntax

 _expression_. **Caption**

 _expression_A variable that represents a  **CheckBox** object.


## Remarks

The default caption for a control is a unique name based on the type of control. For example, CommandButton1 is the default caption for the first command button in a form.

If a control's caption is too long, the caption is truncated. If a form's caption is too long for the title bar, the title is displayed with an ellipsis.

The  **[ForeColor](checkbox-forecolor-property-outlook-forms-script.md)** property of the control determines the color of the text in the caption.

Setting  **[AutoSize](checkbox-autosize-property-outlook-forms-script.md)** to **True** automatically adjusts the size of the control to frame the entire caption.


