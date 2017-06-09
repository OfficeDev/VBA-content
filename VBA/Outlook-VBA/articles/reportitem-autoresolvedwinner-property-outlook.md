---
title: ReportItem.AutoResolvedWinner Property (Outlook)
keywords: vbaol11.chm1678
f1_keywords:
- vbaol11.chm1678
ms.prod: outlook
api_name:
- Outlook.ReportItem.AutoResolvedWinner
ms.assetid: 55f74600-8058-b7cc-33c3-e5b80cef255a
ms.date: 06/08/2017
---


# ReportItem.AutoResolvedWinner Property (Outlook)

Returns a  **Boolean** that determines if the item is a winner of an automatic conflict resolution. Read-only.


## Syntax

 _expression_ . **AutoResolvedWinner**

 _expression_ A variable that represents a **ReportItem** object.


## Remarks

A value of  **False** does not necessarily indicate that the item is a loser of an automatic conflict resolution. The item could be in conflict with another item.

If an item has  **[Conflicts.Count](conflicts-count-property-outlook.md)** of its **[ReportItem.Conflicts](reportitem-conflicts-property-outlook.md)** property greater than zero and if its **AutoResolvedWinner** property is **True** , it is a winner of an automatic conflict resolution. On the other hand, if the item is in conflict and has its **AutoResolvedWinner** property as **False** , it is a loser in an automatic conflict resolution.


## See also


#### Concepts


[ReportItem Object](reportitem-object-outlook.md)

