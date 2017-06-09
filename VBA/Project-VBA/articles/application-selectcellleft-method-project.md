---
title: Application.SelectCellLeft Method (Project)
keywords: vbapj.chm2047
f1_keywords:
- vbapj.chm2047
ms.prod: project-server
api_name:
- Project.Application.SelectCellLeft
ms.assetid: 39bcb2db-cf65-0dc4-2594-9b3c58c4c7c9
ms.date: 06/08/2017
---


# Application.SelectCellLeft Method (Project)

Selects cells to the left of the current selection.


## Syntax

 _expression_. **SelectCellLeft**( ** _NumCells_**, ** _Extend_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _NumCells_|Optional|**Long**|The number of cells to select to the left of the current selection. The default value is 1.|
| _Extend_|Optional|**Boolean**|**True** if the current selection is extended to the specified cell. The default value is **False**.|

### Return Value

 **Boolean**


## Remarks

The  **SelectCellLeft** method is not available when the Calendar, Network Diagram, or Resource Graph is the active view.


