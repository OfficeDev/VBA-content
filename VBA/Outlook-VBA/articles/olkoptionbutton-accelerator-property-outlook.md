---
title: OlkOptionButton.Accelerator Property (Outlook)
keywords: vbaol11.chm1000164
f1_keywords:
- vbaol11.chm1000164
ms.prod: outlook
api_name:
- Outlook.OlkOptionButton.Accelerator
ms.assetid: f1b21d0d-b039-b37b-5f60-4d5acbeaf508
ms.date: 06/08/2017
---


# OlkOptionButton.Accelerator Property (Outlook)

Returns or sets a  **String** value that represents the accelerator or hot key for the control. Read/write.


## Syntax

 _expression_ . **Accelerator**

 _expression_ A variable that represents an **OlkOptionButton** object.


## Remarks

The default value is an empty string, meaning there is no hot key. If the property is set with a string of more than one character, the value will be truncated to the first character. 

You cannot use digits in an accelerator.


## See also


#### Concepts


[OlkOptionButton Object](olkoptionbutton-object-outlook.md)

