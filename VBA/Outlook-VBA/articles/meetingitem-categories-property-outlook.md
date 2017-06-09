---
title: MeetingItem.Categories Property (Outlook)
keywords: vbaol11.chm1406
f1_keywords:
- vbaol11.chm1406
ms.prod: outlook
api_name:
- Outlook.MeetingItem.Categories
ms.assetid: ae4a9569-afb6-a7d7-2cbb-351141f99588
ms.date: 06/08/2017
---


# MeetingItem.Categories Property (Outlook)

Returns or sets a  **String** representing the categories assigned to the Outlook item. Read/write.


## Syntax

 _expression_ . **Categories**

 _expression_ A variable that represents a **MeetingItem** object.


## Remarks

 **Categories** is a delimited string of category names that have been assigned to an Outlook item. This property uses the character specified in the value name, **sList** , under **HKEY_CURRENT_USER\Control Panel\International** in the Windows registry, as the delimiter for multiple categories. To convert the string of category names to an array of category names, use the Microsoft Visual Basic function **Split** .


## See also


#### Concepts


[MeetingItem Object](meetingitem-object-outlook.md)

