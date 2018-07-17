---
title: SelectNamesDialog.BccLabel Property (Outlook)
keywords: vbaol11.chm828
f1_keywords:
- vbaol11.chm828
ms.prod: outlook
api_name:
- Outlook.SelectNamesDialog.BccLabel
ms.assetid: 9c826c3e-c7d3-6fd0-f900-24ba31925681
ms.date: 06/08/2017
---


# SelectNamesDialog.BccLabel Property (Outlook)

Returns or sets a  **String** for the text that appears on the **Bcc** command button on the **Select Names** dialog box. Read/write.


## Syntax

 _expression_ . **BccLabel**

 _expression_ A variable that represents a **SelectNamesDialog** object.


## Remarks

To provide an accelerator key for the recipient edit boxes, include an ampersand (&;) character in the label argument string, immediately before the character that serves as the access key. For example, if  **BccLabel** is the string "Local &;Attendees", users can press **ALT+A** to move the focus to the first recipient edit box.

If  **BccLabel** is not set, then the default value will be the localized string for "Bcc". If the **BccLabel** is set to an empty string, then the **Bcc** command button shows **-&gt;**. If the  **BccLabel** property contains more than 32 characters (including the ampersand (&;) access key), only the first 32 characters will be displayed in the command button.


## See also


#### Concepts


[SelectNamesDialog Object](selectnamesdialog-object-outlook.md)

