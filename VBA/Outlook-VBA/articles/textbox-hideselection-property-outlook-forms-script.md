---
title: TextBox.HideSelection Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 7d59098a-88c3-8086-f8ee-1d9a090865e8
ms.date: 06/08/2017
---


# TextBox.HideSelection Property (Outlook Forms Script)

Returns or sets a  **Boolean** that specifies whether selected text remains highlighted when a control does not have the focus. Read/write.


## Syntax

 _expression_. **HideSelection**

 _expression_A variable that represents a  **TextBox** object.


## Remarks

 **True** if selected text is not highlighted unless the control has the focus (default). **False** if selected text always appears highlighted.

You can use the  **HideSelection** property to maintain highlighted text when another form or a dialog box receives the focus, such as in a spell-checking procedure.


