---
title: Application.SelectEnd Method (Project)
keywords: vbapj.chm2042
f1_keywords:
- vbapj.chm2042
ms.prod: project-server
api_name:
- Project.Application.SelectEnd
ms.assetid: c1d050e7-739d-8a4f-01da-b8c093836733
ms.date: 06/08/2017
---


# Application.SelectEnd Method (Project)

Selects the last cell in the active table that contains a resource or task.


## Syntax

 _expression_. **SelectEnd**( ** _Extend_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Extend_|Optional|**Boolean**|**True** if the current selection is extended to the last cell. If the active view is the Network Diagram or Resource Graph, **Extend** is ignored. The default value is **False**.|

### Return Value

 **Boolean**


## Remarks

In the Resource Graph,  **SelectEnd** selects the resource with the highest identification number. In the Network Diagram, **SelectEnd** selects the box closest to the lower-right corner of the view. The **SelectEnd** method is not available when the Calendar view is active.


