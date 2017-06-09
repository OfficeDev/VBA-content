---
title: Selection.ItemStatus Property (Visio)
keywords: vis_sdr.chm11113780
f1_keywords:
- vis_sdr.chm11113780
ms.prod: visio
api_name:
- Visio.Selection.ItemStatus
ms.assetid: 2dcd9875-222d-fdb9-c2be-1a1df4ee86e7
ms.date: 06/08/2017
---


# Selection.ItemStatus Property (Visio)

Indicates if an item in a  **Selection** object is subselected, if the group to which it belongs is selected, or if it is the primary item. Read-only.


## Syntax

 _expression_ . **ItemStatus**( **_Index_** )

 _expression_ A variable that represents a **Selection** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|Index of the item for which you want to retrieve the status.|

### Return Value

Integer


## Remarks

The  **ItemStatus** property reports a combination of the following values.



|**Constant **|**Value **|**Description **|
|:-----|:-----|:-----|
| **visSelIsPrimaryItem**|&;H1 |The item is the primary item. |
| **visSelIsSubItem**|&;H2 |The item is a subselected item. |
| **visSelIsSuperItem**|&;H4 |The item is a superselected item. |

