---
title: NoteItem.Categories Property (Outlook)
keywords: vbaol11.chm1478
f1_keywords:
- vbaol11.chm1478
ms.prod: outlook
api_name:
- Outlook.NoteItem.Categories
ms.assetid: fd4d258e-fa20-0bdb-a701-8f3c557f0f8a
ms.date: 06/08/2017
---


# NoteItem.Categories Property (Outlook)

Returns or sets a  **String** representing the categories assigned to the Outlook item. Read/write.


## Syntax

 _expression_ . **Categories**

 _expression_ A variable that represents a **NoteItem** object.


## Remarks

 **Categories** is a delimited string of category names that have been assigned to an Outlook item. This property uses the character specified in the value name, **sList** , under **HKEY_CURRENT_USER\Control Panel\International** in the Windows registry, as the delimiter for multiple categories. To convert the string of category names to an array of category names, use the Microsoft Visual Basic function **Split** .


## See also


#### Concepts


[NoteItem Object](noteitem-object-outlook.md)

