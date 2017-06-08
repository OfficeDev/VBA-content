---
title: Document.RenameCurrentScope Method (Visio)
keywords: vis_sdr.chm10550185
f1_keywords:
- vis_sdr.chm10550185
ms.prod: visio
api_name:
- Visio.Document.RenameCurrentScope
ms.assetid: 08aff947-e876-29b8-e910-e2a3b42e5d0e
ms.date: 06/08/2017
---


# Document.RenameCurrentScope Method (Visio)

Renames the top-level open undo scope.


## Syntax

 _expression_ . **RenameCurrentScope**( **_bstrScopeName_** )

 _expression_ A variable that represents a **Document** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _bstrScopeName_|Required| **String**|The new name of the undo scope.|

### Return Value

Nothing


## Remarks

The new name assigned to the undo scope appears on the  **Undo** menu as the item name. If there is no open undo scope, the **RenameCurrentScope** method raises an exception.


