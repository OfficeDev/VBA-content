---
title: TaskRequestUpdateItem.AutoResolvedWinner Property (Outlook)
keywords: vbaol11.chm1962
f1_keywords:
- vbaol11.chm1962
ms.prod: outlook
api_name:
- Outlook.TaskRequestUpdateItem.AutoResolvedWinner
ms.assetid: 06262c98-77ca-623a-1bb4-619a2095193e
ms.date: 06/08/2017
---


# TaskRequestUpdateItem.AutoResolvedWinner Property (Outlook)

Returns a  **Boolean** that determines if the item is a winner of an automatic conflict resolution. Read-only.


## Syntax

 _expression_ . **AutoResolvedWinner**

 _expression_ A variable that represents a **TaskRequestUpdateItem** object.


## Remarks

A value of  **False** does not necessarily indicate that the item is a loser of an automatic conflict resolution. The item could be in conflict with another item.

If an item has  **[Conflicts.Count](conflicts-count-property-outlook.md)** of its **[TaskRequestUpdateItem.Conflicts](taskrequestupdateitem-conflicts-property-outlook.md)** property greater than zero and if its **AutoResolvedWinner** property is **True** , it is a winner of an automatic conflict resolution. On the other hand, if the item is in conflict and has its **AutoResolvedWinner** property as **False** , it is a loser in an automatic conflict resolution.


## See also


#### Concepts


[TaskRequestUpdateItem Object](taskrequestupdateitem-object-outlook.md)

