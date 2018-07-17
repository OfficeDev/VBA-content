---
title: DistListItem.IsConflict Property (Outlook)
keywords: vbaol11.chm1163
f1_keywords:
- vbaol11.chm1163
ms.prod: outlook
api_name:
- Outlook.DistListItem.IsConflict
ms.assetid: 3c1417a8-6609-c715-04f1-625ea733134c
ms.date: 06/08/2017
---


# DistListItem.IsConflict Property (Outlook)

Returns a  **Boolean** that determines if the item is in conflict. Read-only.


## Syntax

 _expression_ . **IsConflict**

 _expression_ A variable that represents a **DistListItem** object.


## Remarks

Whether or not an item is in conflict is determined by the state of the application. For example, when a user is offline and tries to access an online folder the action will fail. In this scenario, the  **IsConflict** property will return **True** .

If  **True** , the specified item is in conflict.


## See also


#### Concepts


[DistListItem Object](distlistitem-object-outlook.md)

