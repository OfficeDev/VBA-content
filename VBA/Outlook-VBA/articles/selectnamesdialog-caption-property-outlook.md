---
title: SelectNamesDialog.Caption Property (Outlook)
keywords: vbaol11.chm825
f1_keywords:
- vbaol11.chm825
ms.prod: outlook
api_name:
- Outlook.SelectNamesDialog.Caption
ms.assetid: a728bcb5-8eee-8f77-76d7-4c15d53d79e2
ms.date: 06/08/2017
---


# SelectNamesDialog.Caption Property (Outlook)

Returns or sets a  **String** value that represents the title for the **Select Names** dialog box. Read/write.


## Syntax

 _expression_ . **Caption**

 _expression_ A variable that represents a **SelectNamesDialog** object.


## Remarks

If you do not set  **Caption** , the title of the dialog box will be **Select Names** or the localized equivalent. If you set **Caption** to an empty string, the dialog box caption will be an empty string.

Setting the  **Caption** to a long string (for example, 300 characters) will cause the caption to be truncated, and will not cause the width of the **Select Names** dialog to change.


## See also


#### Concepts


[SelectNamesDialog Object](selectnamesdialog-object-outlook.md)

