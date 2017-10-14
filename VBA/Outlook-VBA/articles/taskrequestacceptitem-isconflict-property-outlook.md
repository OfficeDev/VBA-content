---
title: TaskRequestAcceptItem.IsConflict Property (Outlook)
keywords: vbaol11.chm1814
f1_keywords:
- vbaol11.chm1814
ms.prod: outlook
api_name:
- Outlook.TaskRequestAcceptItem.IsConflict
ms.assetid: e6e362d2-18c4-ca68-8c8f-fbd11482e597
ms.date: 06/08/2017
---


# TaskRequestAcceptItem.IsConflict Property (Outlook)

Returns a  **Boolean** that determines if the item is in conflict. Read-only.


## Syntax

 _expression_ . **IsConflict**

 _expression_ A variable that represents a **TaskRequestAcceptItem** object.


## Remarks

Whether or not an item is in conflict is determined by the state of the application. For example, when a user is offline and tries to access an online folder the action will fail. In this scenario, the  **IsConflict** property will return **True** .

If  **True** , the specified item is in conflict.


## See also


#### Concepts


[TaskRequestAcceptItem Object](taskrequestacceptitem-object-outlook.md)

