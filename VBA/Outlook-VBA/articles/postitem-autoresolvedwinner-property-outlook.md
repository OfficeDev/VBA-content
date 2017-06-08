---
title: PostItem.AutoResolvedWinner Property (Outlook)
keywords: vbaol11.chm1565
f1_keywords:
- vbaol11.chm1565
ms.prod: outlook
api_name:
- Outlook.PostItem.AutoResolvedWinner
ms.assetid: 4ad920e3-0487-f63c-1af9-2ac623d2487b
ms.date: 06/08/2017
---


# PostItem.AutoResolvedWinner Property (Outlook)

Returns a  **Boolean** that determines if the item is a winner of an automatic conflict resolution. Read-only.


## Syntax

 _expression_ . **AutoResolvedWinner**

 _expression_ A variable that represents a **PostItem** object.


## Remarks

A value of  **False** does not necessarily indicate that the item is a loser of an automatic conflict resolution. The item could be in conflict with another item.

If an item has  **[Conflicts.Count](conflicts-count-property-outlook.md)** of its **[PostItem.Conflicts](postitem-conflicts-property-outlook.md)** property greater than zero and if its **AutoResolvedWinner** property is **True** , it is a winner of an automatic conflict resolution. On the other hand, if the item is in conflict and has its **AutoResolvedWinner** property as **False** , it is a loser in an automatic conflict resolution.


## See also


#### Concepts


[PostItem Object](postitem-object-outlook.md)

