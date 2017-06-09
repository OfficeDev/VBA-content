---
title: TaskRequestItem.Categories Property (Outlook)
keywords: vbaol11.chm1876
f1_keywords:
- vbaol11.chm1876
ms.prod: outlook
api_name:
- Outlook.TaskRequestItem.Categories
ms.assetid: 12f9bc19-de45-47d4-ca2a-2beb1ddadefa
ms.date: 06/08/2017
---


# TaskRequestItem.Categories Property (Outlook)

Returns or sets a  **String** representing the categories assigned to the Outlook item. Read/write.


## Syntax

 _expression_ . **Categories**

 _expression_ A variable that represents a **TaskRequestItem** object.


## Remarks

 **Categories** is a delimited string of category names that have been assigned to an Outlook item. This property uses the character specified in the value name, **sList** , under **HKEY_CURRENT_USER\Control Panel\International** in the Windows registry, as the delimiter for multiple categories. To convert the string of category names to an array of category names, use the Microsoft Visual Basic function **Split** .


## See also


#### Concepts


[TaskRequestItem Object](taskrequestitem-object-outlook.md)

