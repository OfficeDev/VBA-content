---
title: TaskRequestUpdateItem.IsConflict Property (Outlook)
keywords: vbaol11.chm1961
f1_keywords:
- vbaol11.chm1961
ms.prod: outlook
api_name:
- Outlook.TaskRequestUpdateItem.IsConflict
ms.assetid: c46f3c3a-57b0-facd-4090-7568f1b78667
ms.date: 06/08/2017
---


# TaskRequestUpdateItem.IsConflict Property (Outlook)

Returns a  **Boolean** that determines if the item is in conflict. Read-only.


## Syntax

 _expression_ . **IsConflict**

 _expression_ A variable that represents a **TaskRequestUpdateItem** object.


## Remarks

Whether or not an item is in conflict is determined by the state of the application. For example, when a user is offline and tries to access an online folder the action will fail. In this scenario, the  **IsConflict** property will return **True** .

If  **True** , the specified item is in conflict.


## See also


#### Concepts


[TaskRequestUpdateItem Object](taskrequestupdateitem-object-outlook.md)

