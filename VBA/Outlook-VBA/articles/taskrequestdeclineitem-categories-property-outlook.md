---
title: TaskRequestDeclineItem.Categories Property (Outlook)
keywords: vbaol11.chm1827
f1_keywords:
- vbaol11.chm1827
ms.prod: outlook
api_name:
- Outlook.TaskRequestDeclineItem.Categories
ms.assetid: 11ac178b-c43d-c6ac-f4d9-2b016b2f3793
ms.date: 06/08/2017
---


# TaskRequestDeclineItem.Categories Property (Outlook)

Returns or sets a  **String** representing the categories assigned to the Outlook item. Read/write.


## Syntax

 _expression_ . **Categories**

 _expression_ A variable that represents a **TaskRequestDeclineItem** object.


## Remarks

 **Categories** is a delimited string of category names that have been assigned to an Outlook item. This property uses the character specified in the value name, **sList** , under **HKEY_CURRENT_USER\Control Panel\International** in the Windows registry, as the delimiter for multiple categories. To convert the string of category names to an array of category names, use the Microsoft Visual Basic function **Split** .


## See also


#### Concepts


[TaskRequestDeclineItem Object](taskrequestdeclineitem-object-outlook.md)

