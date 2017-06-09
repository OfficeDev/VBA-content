---
title: DistListItem.AutoResolvedWinner Property (Outlook)
keywords: vbaol11.chm1164
f1_keywords:
- vbaol11.chm1164
ms.prod: outlook
api_name:
- Outlook.DistListItem.AutoResolvedWinner
ms.assetid: cb43f885-07b0-aa7c-a055-7eb8027ee766
ms.date: 06/08/2017
---


# DistListItem.AutoResolvedWinner Property (Outlook)

Returns a  **Boolean** that determines if the item is a winner of an automatic conflict resolution. Read-only.


## Syntax

 _expression_ . **AutoResolvedWinner**

 _expression_ A variable that represents a **DistListItem** object.


## Remarks

A value of  **False** does not necessarily indicate that the item is a loser of an automatic conflict resolution. The item could be in conflict with another item.

If an item has  **[Conflicts.Count](conflicts-count-property-outlook.md)** of its **[DistListItem.Conflicts](distlistitem-conflicts-property-outlook.md)** property greater than zero and if its **AutoResolvedWinner** property is **True** , it is a winner of an automatic conflict resolution. On the other hand, if the item is in conflict and has its **AutoResolvedWinner** property as **False** , it is a loser in an automatic conflict resolution.


## See also


#### Concepts


[DistListItem Object](distlistitem-object-outlook.md)

