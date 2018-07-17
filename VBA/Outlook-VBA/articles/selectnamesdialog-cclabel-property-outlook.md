---
title: SelectNamesDialog.CcLabel Property (Outlook)
keywords: vbaol11.chm829
f1_keywords:
- vbaol11.chm829
ms.prod: outlook
api_name:
- Outlook.SelectNamesDialog.CcLabel
ms.assetid: b28def6f-725c-ba65-cf7f-4abbc7ba3cb8
ms.date: 06/08/2017
---


# SelectNamesDialog.CcLabel Property (Outlook)

Returns or sets a  **String** for the text that appears on the **Cc** command button on the **Select Names** dialog box. Read/write.


## Syntax

 _expression_ . **CcLabel**

 _expression_ A variable that represents a **SelectNamesDialog** object.


## Remarks

To provide an accelerator key for the recipient edit boxes, include an ampersand (&;) character in the label argument string, immediately before the character that serves as the access key. For example, if  **CcLabel** is the string "Local &;Attendees", users can press **ALT+A** to move the focus to the first recipient edit box.

If  **CcLabel** is not set, then the default value will be the localized string for "Cc". If the **CcLabel** is set to an empty string, then the **Cc** command button shows **-&gt;**. If the  **CcLabel** property contains more than 32 characters (including the ampersand (&;) access key), only the first 32 characters will be displayed in the command button.


## See also


#### Concepts


[SelectNamesDialog Object](selectnamesdialog-object-outlook.md)

