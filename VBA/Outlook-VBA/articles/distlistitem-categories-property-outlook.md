---
title: DistListItem.Categories Property (Outlook)
keywords: vbaol11.chm1118
f1_keywords:
- vbaol11.chm1118
ms.prod: outlook
api_name:
- Outlook.DistListItem.Categories
ms.assetid: b608ce9d-8419-cf70-716e-0c4cdca2fa98
ms.date: 06/08/2017
---


# DistListItem.Categories Property (Outlook)

Returns or sets a  **String** representing the categories assigned to the Outlook item. Read/write.


## Syntax

 _expression_ . **Categories**

 _expression_ A variable that represents a **DistListItem** object.


## Remarks

 **Categories** is a delimited string of category names that have been assigned to an Outlook item. This property uses the character specified in the value name, **sList** , under **HKEY_CURRENT_USER\Control Panel\International** in the Windows registry, as the delimiter for multiple categories. To convert the string of category names to an array of category names, use the Microsoft Visual Basic function **Split** .


## See also


#### Concepts


[DistListItem Object](distlistitem-object-outlook.md)

