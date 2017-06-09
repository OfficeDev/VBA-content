---
title: Application.RenameCurrentScope Method (Visio)
keywords: vis_sdr.chm10050815
f1_keywords:
- vis_sdr.chm10050815
ms.prod: visio
api_name:
- Visio.Application.RenameCurrentScope
ms.assetid: 0ccd9c6b-704c-b956-5ea9-4f1ed01baee7
ms.date: 06/08/2017
---


# Application.RenameCurrentScope Method (Visio)

Renames the top-level open undo scope.


## Syntax

 _expression_ . **RenameCurrentScope**( **_bstrScopeName_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _bstrScopeName_|Required| **String**|The new name of the undo scope.|

### Return Value

Nothing


## Remarks

The new name assigned to the undo scope appears on the  **Undo** menu as the item name. If there is no open undo scope, the **RenameCurrentScope** method raises an exception.


