---
title: Application.OnUndoOrRedo Event (Project)
keywords: vbapj.chm131132
f1_keywords:
- vbapj.chm131132
ms.prod: project-server
api_name:
- Project.Application.OnUndoOrRedo
ms.assetid: 7f60e893-81d0-1b2f-c5f5-ec1451633fa7
ms.date: 06/08/2017
---


# Application.OnUndoOrRedo Event (Project)

Occurs when a transaction is undone or redone.


## Syntax

 _expression_. **OnUndoOrRedo**( ** _bstrLabel_**, ** _bstrGUID_**, ** _fUndo_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _bstrLabel_|Required|**String**|Label of the transaction just undone or redone.|
| _bstrGUID_|Required|**String**|GUID of the transaction or NULL.|
| _fUndo_|Required|**Boolean**|**True** if the transaction was undone or **False** if it was redone.|

### Return Value

nothing


## Remarks

You can use the  **OnUndoOrRedo** event to manage undo or redo actions that are specified by the global **OpenUndoTransaction** and **CloseUndoTransaction** methods.


