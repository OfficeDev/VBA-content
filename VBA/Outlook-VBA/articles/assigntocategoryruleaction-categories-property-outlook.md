---
title: AssignToCategoryRuleAction.Categories Property (Outlook)
keywords: vbaol11.chm2269
f1_keywords:
- vbaol11.chm2269
ms.prod: outlook
api_name:
- Outlook.AssignToCategoryRuleAction.Categories
ms.assetid: 92e849e3-4d5a-a11b-3c32-6214a15a90df
ms.date: 06/08/2017
---


# AssignToCategoryRuleAction.Categories Property (Outlook)

Returns or sets an array of strings representing the categories assigned to the message. Read/write.


## Syntax

 _expression_ . **Categories**

 _expression_ A variable that represents an **AssignToCategoryRuleAction** object.


## Remarks

You can assign an array with one element for a single category or an array of strings for multiple categories. Outlook does not check to determine if the  **Categories** property contains category names that are in the master category list.

This property uses the character specified in the value name,  **sList** , under **HKEY_CURRENT_USER\Control Panel\International** in the Windows registry, as the delimiter for multiple categories.


## See also


#### Concepts


[AssignToCategoryRuleAction Object](assigntocategoryruleaction-object-outlook.md)

