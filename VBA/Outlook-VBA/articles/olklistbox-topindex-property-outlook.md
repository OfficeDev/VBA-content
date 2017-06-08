---
title: OlkListBox.TopIndex Property (Outlook)
keywords: vbaol11.chm1000264
f1_keywords:
- vbaol11.chm1000264
ms.prod: outlook
api_name:
- Outlook.OlkListBox.TopIndex
ms.assetid: 8d024de7-4135-4957-4d84-1b0199219f8f
ms.date: 06/08/2017
---


# OlkListBox.TopIndex Property (Outlook)

Returns or sets a  **Long** that represents the index of the item at the top of the displayed portion of the list. Read/write.


## Syntax

 _expression_ . **TopIndex**

 _expression_ A variable that represents an **OlkListBox** object.


## Remarks

As the list scrolls, the item at the top of the list will change, and the value of this property will change to reflect the item currently displayed at the top of the list.

The index value is zero-based. The default value is -1, indicating that no special ordering should be applied.


## See also


#### Concepts


[OlkListBox Object](olklistbox-object-outlook.md)

