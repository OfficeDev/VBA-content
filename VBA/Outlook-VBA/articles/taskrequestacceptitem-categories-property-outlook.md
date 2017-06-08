---
title: TaskRequestAcceptItem.Categories Property (Outlook)
keywords: vbaol11.chm1778
f1_keywords:
- vbaol11.chm1778
ms.prod: outlook
api_name:
- Outlook.TaskRequestAcceptItem.Categories
ms.assetid: 18b34d77-3479-08b3-d031-4732fb7657f1
ms.date: 06/08/2017
---


# TaskRequestAcceptItem.Categories Property (Outlook)

Returns or sets a  **String** representing the categories assigned to the Outlook item. Read/write.


## Syntax

 _expression_ . **Categories**

 _expression_ A variable that represents a **TaskRequestAcceptItem** object.


## Remarks

 **Categories** is a delimited string of category names that have been assigned to an Outlook item. This property uses the character specified in the value name, **sList** , under **HKEY_CURRENT_USER\Control Panel\International** in the Windows registry, as the delimiter for multiple categories. To convert the string of category names to an array of category names, use the Microsoft Visual Basic function **Split** .


## See also


#### Concepts


[TaskRequestAcceptItem Object](taskrequestacceptitem-object-outlook.md)

