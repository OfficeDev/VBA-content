---
title: Application.SelectCellRight Method (Project)
keywords: vbapj.chm2048
f1_keywords:
- vbapj.chm2048
ms.prod: project-server
api_name:
- Project.Application.SelectCellRight
ms.assetid: 3753531a-5459-eb25-a9b9-2e9f748a0518
ms.date: 06/08/2017
---


# Application.SelectCellRight Method (Project)

Selects cells to the right of the current selection.


## Syntax

 _expression_. **SelectCellRight**( ** _NumCells_**, ** _Extend_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _NumCells_|Optional|**Long**|The number of cells to select to the right of the current selection. The default value is 1.|
| _Extend_|Optional|**Boolean**|**True** if the current selection is extended to the specified cell. The default value is **False**.|

### Return Value

 **Boolean**


## Remarks

The  **SelectCellRight** method is not available when the Calendar, Network Diagram, or Resource Graph is the active view.


