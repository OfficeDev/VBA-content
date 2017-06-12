---
title: Application.SelectCellUp Method (Project)
keywords: vbapj.chm2049
f1_keywords:
- vbapj.chm2049
ms.prod: project-server
api_name:
- Project.Application.SelectCellUp
ms.assetid: d2e2aecc-0a05-7dd5-23da-a47ffe161028
ms.date: 06/08/2017
---


# Application.SelectCellUp Method (Project)

Selects cells upward from the current selection.


## Syntax

 _expression_. **SelectCellUp**( ** _NumCells_**, ** _Extend_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _NumCells_|Optional|**Long**|The number of cells to select upward from the current selection. The default value is 1.|
| _Extend_|Optional|**Boolean**|**True** if the current selection is extended to the specified cell. The default value is **False**.|

### Return Value

 **Boolean**


## Remarks

The  **SelectCellUp** method is not available when the Calendar, Network Diagram, or Resource Graph is the active view.


