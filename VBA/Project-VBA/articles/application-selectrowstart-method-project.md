---
title: Application.SelectRowStart Method (Project)
keywords: vbapj.chm2043
f1_keywords:
- vbapj.chm2043
ms.prod: project-server
api_name:
- Project.Application.SelectRowStart
ms.assetid: cbb2c5a8-edbb-5d5e-e4ef-5a952db769c3
ms.date: 06/08/2017
---


# Application.SelectRowStart Method (Project)

Selects the first cell in the row containing the active cell.


## Syntax

 _expression_. **SelectRowStart**( ** _Extend_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Extend_|Optional|**Boolean**|**True** if the current selection is extended to the first cell. The default value is **False**.|

### Return Value

 **Boolean**


## Remarks

The  **SelectRowStart** method is only available when the Gantt Chart, Task Sheet, Task Usage view, Resource Sheet, or Resource Usage view is the active view.


