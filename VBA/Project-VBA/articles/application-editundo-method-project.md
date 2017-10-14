---
title: Application.EditUndo Method (Project)
keywords: vbapj.chm201
f1_keywords:
- vbapj.chm201
ms.prod: project-server
api_name:
- Project.Application.EditUndo
ms.assetid: f13ce3a1-f8f2-8b00-d870-6e30f6b772f5
ms.date: 06/08/2017
---


# Application.EditUndo Method (Project)

Cancels the last user-interface action.


## Syntax

 _expression_. **EditUndo**( ** _fUndo_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _fUndo_|Optional|**Integer**|Specifies the number of actions to undo. If the total number of actions is less than fUndo,  **EditUndo** undoes all actions.|

### Return Value

 **Boolean**


