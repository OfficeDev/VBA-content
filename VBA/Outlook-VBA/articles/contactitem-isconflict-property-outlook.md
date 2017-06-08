---
title: ContactItem.IsConflict Property (Outlook)
keywords: vbaol11.chm1087
f1_keywords:
- vbaol11.chm1087
ms.prod: outlook
api_name:
- Outlook.ContactItem.IsConflict
ms.assetid: 35ff3a52-2d2a-458f-3e16-4a8f674bb0fa
ms.date: 06/08/2017
---


# ContactItem.IsConflict Property (Outlook)

Returns a  **Boolean** that determines if the item is in conflict. Read-only.


## Syntax

 _expression_ . **IsConflict**

 _expression_ A variable that represents a **ContactItem** object.


## Remarks

Whether or not an item is in conflict is determined by the state of the application. For example, when a user is offline and tries to access an online folder the action will fail. In this scenario, the  **IsConflict** property will return **True** .

If  **True** , the specified item is in conflict.


## See also


#### Concepts


[ContactItem Object](contactitem-object-outlook.md)

