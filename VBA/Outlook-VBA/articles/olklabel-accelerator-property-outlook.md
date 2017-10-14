---
title: OlkLabel.Accelerator Property (Outlook)
keywords: vbaol11.chm1000086
f1_keywords:
- vbaol11.chm1000086
ms.prod: outlook
api_name:
- Outlook.OlkLabel.Accelerator
ms.assetid: 7d461585-5aa1-81ab-8cec-5e25795e9bea
ms.date: 06/08/2017
---


# OlkLabel.Accelerator Property (Outlook)

Returns or sets a  **String** value that represents the accelerator or hot key for the control. Read/write.


## Syntax

 _expression_ . **Accelerator**

 _expression_ A variable that represents an **OlkLabel** object.


## Remarks

The default value is an empty string, meaning there is no hot key. If the property is set with a string of more than one character, the value will be truncated to the first character. 

You cannot use digits in an accelerator.

When the accelerator key for a label is pressed, the next control in the tab order receives the focus, not the label control.


## See also


#### Concepts


[OlkLabel Object](olklabel-object-outlook.md)

