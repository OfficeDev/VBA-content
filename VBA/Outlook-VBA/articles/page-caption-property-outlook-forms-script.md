---
title: Page.Caption Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 878ed59e-8aa9-ec07-487a-47706d5337f4
ms.date: 06/08/2017
---


# Page.Caption Property (Outlook Forms Script)

Returns or sets a  **String** that specifies the text that appears on the page. Read/write.


## Syntax

 _expression_. **Caption**

 _expression_A variable that represents a  **Page** object.


## Remarks

The default caption for an object is a unique name based on the type of object. For example, CommandButton1 is the default caption for the first command button in a form.

If an object's caption is too long, the caption is truncated. If a form's caption is too long for the title bar, the title is displayed with an ellipsis.


