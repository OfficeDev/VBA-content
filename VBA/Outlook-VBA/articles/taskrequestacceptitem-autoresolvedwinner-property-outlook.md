---
title: TaskRequestAcceptItem.AutoResolvedWinner Property (Outlook)
keywords: vbaol11.chm1815
f1_keywords:
- vbaol11.chm1815
ms.prod: outlook
api_name:
- Outlook.TaskRequestAcceptItem.AutoResolvedWinner
ms.assetid: 2d49d695-865f-d60d-6d0f-ced8268d89d6
ms.date: 06/08/2017
---


# TaskRequestAcceptItem.AutoResolvedWinner Property (Outlook)

Returns a  **Boolean** that determines if the item is a winner of an automatic conflict resolution. Read-only.


## Syntax

 _expression_ . **AutoResolvedWinner**

 _expression_ A variable that represents a **TaskRequestAcceptItem** object.


## Remarks

A value of  **False** does not necessarily indicate that the item is a loser of an automatic conflict resolution. The item could be in conflict with another item.

If an item has  **[Conflicts.Count](conflicts-count-property-outlook.md)** of its **[TaskRequestAcceptItem.Conflicts](taskrequestacceptitem-conflicts-property-outlook.md)** property greater than zero and if its **AutoResolvedWinner** property is **True** , it is a winner of an automatic conflict resolution. On the other hand, if the item is in conflict and has its **AutoResolvedWinner** property as **False** , it is a loser in an automatic conflict resolution.


## See also


#### Concepts


[TaskRequestAcceptItem Object](taskrequestacceptitem-object-outlook.md)

