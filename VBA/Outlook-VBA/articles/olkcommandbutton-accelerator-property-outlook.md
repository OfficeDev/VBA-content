---
title: OlkCommandButton.Accelerator Property (Outlook)
keywords: vbaol11.chm1000109
f1_keywords:
- vbaol11.chm1000109
ms.prod: outlook
api_name:
- Outlook.OlkCommandButton.Accelerator
ms.assetid: c66b26b7-f17f-ce2f-c871-49f0eac12913
ms.date: 06/08/2017
---


# OlkCommandButton.Accelerator Property (Outlook)

Returns or sets a  **String** value that represents the accelerator or hot key for the control. Read/write.


## Syntax

 _expression_ . **Accelerator**

 _expression_ A variable that represents an **OlkCommandButton** object.


## Remarks

The default value is an empty string, meaning there is no hot key. If the property is set with a string of more than one character, the value will be truncated to the first character. 

You cannot use digits in an accelerator.


## See also


#### Concepts


[OlkCommandButton Object](olkcommandbutton-object-outlook.md)

