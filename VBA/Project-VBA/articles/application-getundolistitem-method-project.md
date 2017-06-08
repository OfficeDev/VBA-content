---
title: Application.GetUndoListItem Method (Project)
keywords: vbapj.chm131097
f1_keywords:
- vbapj.chm131097
ms.prod: project-server
api_name:
- Project.Application.GetUndoListItem
ms.assetid: e77826ab-118d-2b69-6f99-cb8ce65afb43
ms.date: 06/08/2017
---


# Application.GetUndoListItem Method (Project)

Returns the label of the specified undo list item.


## Syntax

 _expression_. **GetUndoListItem**( ** _ItemIndex_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ItemIndex_|Required|**Long**|Index of the item in the undo list .|

### Return Value

 **String**


## Example

The following example returns the label of the first item in the undo list.


```
GetUndoListItem(1)
```


