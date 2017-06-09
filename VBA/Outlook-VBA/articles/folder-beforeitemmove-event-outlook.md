---
title: Folder.BeforeItemMove Event (Outlook)
keywords: vbaol11.chm2752
f1_keywords:
- vbaol11.chm2752
ms.prod: outlook
api_name:
- Outlook.Folder.BeforeItemMove
ms.assetid: db75bc05-c80e-e6b8-d017-2150bc942712
ms.date: 06/08/2017
---


# Folder.BeforeItemMove Event (Outlook)

Occurs when an item is about to be moved or deleted from a folder, either as a result of user action or through program code. 


## Syntax

 _expression_ . **BeforeItemMove**( **_Item_** , **_MoveTo_** , **_Cancel_** )

 _expression_ A variable that represents a **Folder** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Item_|Required| **Object**|Represents the Outlook item that is to be moved or deleted.|
| _MoveTo_|Required| **[Folder](folder-object-outlook.md)**|Represents the folder to which the item is being moved. |
| _Cancel_|Required| **Boolean**|Set this to  **True** to cancel the move or delete action.|

## Remarks

This event fires when the item is about to be moved to another folder (including the Deleted Items folder) or when the item is about to be permanently deleted. It does not fire during auto-archiving or synchronizing operations.

If the action is a permanent delete, the  _MoveTo_ folder returned in the event will be **Null** ( **Nothing** in Visual Basic).


## See also


#### Concepts


[Folder Object](folder-object-outlook.md)

