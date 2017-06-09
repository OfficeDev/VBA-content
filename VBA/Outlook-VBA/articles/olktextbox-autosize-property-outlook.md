---
title: OlkTextBox.AutoSize Property (Outlook)
keywords: vbaol11.chm1000035
f1_keywords:
- vbaol11.chm1000035
ms.prod: outlook
api_name:
- Outlook.OlkTextBox.AutoSize
ms.assetid: 2445da74-24ff-8f22-a55a-b6f39a79129b
ms.date: 06/08/2017
---


# OlkTextBox.AutoSize Property (Outlook)

Returns or sets a  **Boolean** that automatically sizes the control to display the entire contents. Read/write.


## Syntax

 _expression_ . **AutoSize**

 _expression_ A variable that represents an **OlkTextBox** object.


## Remarks

 The default value for this property is **False** .

For a single-line text box, setting  **AutoSize** to **True** automatically sets the width of the display area to the length of the text in the text box.

For a multiline text box that contains no text, setting  **AutoSize** to **True** automatically displays the text as a column. The width of the text column is set to accommodate the widest letter of that font size. The height of the text column is set to display the entire text of the text box. For a multiline text box that contains text, setting **AutoSize** to **True** automatically extends the text box vertically to display the entire text. The width of the text box does not change.


## See also


#### Concepts


[OlkTextBox Object](olktextbox-object-outlook.md)

