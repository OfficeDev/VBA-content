---
title: NoteItem.IsConflict Property (Outlook)
keywords: vbaol11.chm1504
f1_keywords:
- vbaol11.chm1504
ms.prod: outlook
api_name:
- Outlook.NoteItem.IsConflict
ms.assetid: 5fc4880f-8e96-9993-9b93-341f7a57e420
ms.date: 06/08/2017
---


# NoteItem.IsConflict Property (Outlook)

Returns a  **Boolean** that determines if the item is in conflict. Read-only.


## Syntax

 _expression_ . **IsConflict**

 _expression_ A variable that represents a **NoteItem** object.


## Remarks

Whether or not an item is in conflict is determined by the state of the application. For example, when a user is offline and tries to access an online folder the action will fail. In this scenario, the  **IsConflict** property will return **True** .

If  **True** , the specified item is in conflict.


## See also


#### Concepts


[NoteItem Object](noteitem-object-outlook.md)

