---
title: PostItem.Categories Property (Outlook)
keywords: vbaol11.chm1517
f1_keywords:
- vbaol11.chm1517
ms.prod: outlook
api_name:
- Outlook.PostItem.Categories
ms.assetid: 6ecad980-eedd-b27a-1677-f8027adb6de8
ms.date: 06/08/2017
---


# PostItem.Categories Property (Outlook)

Returns or sets a  **String** representing the categories assigned to the Outlook item. Read/write.


## Syntax

 _expression_ . **Categories**

 _expression_ A variable that represents a **PostItem** object.


## Remarks

 **Categories** is a delimited string of category names that have been assigned to an Outlook item. This property uses the character specified in the value name, **sList** , under **HKEY_CURRENT_USER\Control Panel\International** in the Windows registry, as the delimiter for multiple categories. To convert the string of category names to an array of category names, use the Microsoft Visual Basic function **Split** .


## See also


#### Concepts


[PostItem Object](postitem-object-outlook.md)

