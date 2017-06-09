---
title: ReportItem.Categories Property (Outlook)
keywords: vbaol11.chm1642
f1_keywords:
- vbaol11.chm1642
ms.prod: outlook
api_name:
- Outlook.ReportItem.Categories
ms.assetid: 57983279-5be9-1a08-8a13-d70d5e252699
ms.date: 06/08/2017
---


# ReportItem.Categories Property (Outlook)

Returns or sets a  **String** representing the categories assigned to the Outlook item. Read/write.


## Syntax

 _expression_ . **Categories**

 _expression_ A variable that represents a **ReportItem** object.


## Remarks

 **Categories** is a delimited string of category names that have been assigned to an Outlook item. This property uses the character specified in the value name, **sList** , under **HKEY_CURRENT_USER\Control Panel\International** in the Windows registry, as the delimiter for multiple categories. To convert the string of category names to an array of category names, use the Microsoft Visual Basic function **Split** .


## See also


#### Concepts


[ReportItem Object](reportitem-object-outlook.md)

