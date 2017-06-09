---
title: RemoteItem.IsConflict Property (Outlook)
keywords: vbaol11.chm1629
f1_keywords:
- vbaol11.chm1629
ms.prod: outlook
api_name:
- Outlook.RemoteItem.IsConflict
ms.assetid: 56c3aa72-4ddf-802e-b6ab-7e982a80dc08
ms.date: 06/08/2017
---


# RemoteItem.IsConflict Property (Outlook)

Returns a  **Boolean** that determines if the item is in conflict. Read-only.


## Syntax

 _expression_ . **IsConflict**

 _expression_ A variable that represents a **RemoteItem** object.


## Remarks

Whether or not an item is in conflict is determined by the state of the application. For example, when a user is offline and tries to access an online folder the action will fail. In this scenario, the  **IsConflict** property will return **True** .

If  **True** , the specified item is in conflict.


## See also


#### Concepts


[RemoteItem Object](remoteitem-object-outlook.md)

