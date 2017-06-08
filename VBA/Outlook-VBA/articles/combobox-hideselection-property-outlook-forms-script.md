---
title: ComboBox.HideSelection Property (Outlook Forms Script)
keywords: olfm10.chm2001270
f1_keywords:
- olfm10.chm2001270
ms.prod: outlook
ms.assetid: 4306fbaa-9afa-735a-7195-887977e9ce4c
ms.date: 06/08/2017
---


# ComboBox.HideSelection Property (Outlook Forms Script)

Returns or sets a  **Boolean** that specifies whether selected text remains highlighted when a control does not have the focus. Read/write.


## Syntax

 _expression_. **HideSelection**

 _expression_A variable that represents a  **ComboBox** object.


## Remarks

 **True** if selected text is not highlighted unless the control has the focus (default). **False** if selected text always appears highlighted.

You can use the  **HideSelection** property to maintain highlighted text when another form or a dialog box receives the focus, such as in a spell-checking procedure.


