---
title: Folder.BeforeFolderMove Event (Outlook)
keywords: vbaol11.chm2751
f1_keywords:
- vbaol11.chm2751
ms.prod: outlook
api_name:
- Outlook.Folder.BeforeFolderMove
ms.assetid: c085f0cf-3d91-db84-aab9-18c7b46a04d2
ms.date: 06/08/2017
---


# Folder.BeforeFolderMove Event (Outlook)

Occurs when a folder is about to be moved or deleted, either as a result of user action or through program code. 


## Syntax

 _expression_ . **BeforeFolderMove**( **_MoveTo_** , **_Cancel_** )

 _expression_ A variable that represents a **Folder** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _MoveTo_|Required| **[Folder](folder-object-outlook.md)**|Represents the folder to which the item is being moved. |
| _Cancel_|Required| **Boolean**|Set this to  **True** to cancel the move or delete action.|

## Remarks

This event fires when the folder is about to be moved to another folder (including the Deleted Items folder) or when the folder is about to be permanently deleted. It does not fire during auto-archiving or synchronizing operations.

If the action is a permanent delete, the  _MoveTo_ folder returned in the event will be **Null** ( **Nothing** in Visual Basic).


## See also


#### Concepts


[Folder Object](folder-object-outlook.md)

