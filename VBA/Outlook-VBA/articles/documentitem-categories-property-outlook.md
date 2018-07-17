---
title: DocumentItem.Categories Property (Outlook)
keywords: vbaol11.chm1187
f1_keywords:
- vbaol11.chm1187
ms.prod: outlook
api_name:
- Outlook.DocumentItem.Categories
ms.assetid: 2aa3df17-39f4-6e9c-a32d-5491d17dcb8e
ms.date: 06/08/2017
---


# DocumentItem.Categories Property (Outlook)

Returns or sets a  **String** representing the categories assigned to the Outlook item. Read/write.


## Syntax

 _expression_ . **Categories**

 _expression_ A variable that represents a **DocumentItem** object.


## Remarks

 **Categories** is a delimited string of category names that have been assigned to an Outlook item. This property uses the character specified in the value name, **sList** , under **HKEY_CURRENT_USER\Control Panel\International** in the Windows registry, as the delimiter for multiple categories. To convert the string of category names to an array of category names, use the Microsoft Visual Basic function **Split** .


## See also


#### Concepts


[DocumentItem Object](documentitem-object-outlook.md)

