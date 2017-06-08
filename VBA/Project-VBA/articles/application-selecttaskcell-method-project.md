---
title: Application.SelectTaskCell Method (Project)
keywords: vbapj.chm2068
f1_keywords:
- vbapj.chm2068
ms.prod: project-server
api_name:
- Project.Application.SelectTaskCell
ms.assetid: 824be785-faa8-b274-bc4c-b830f828528d
ms.date: 06/08/2017
---


# Application.SelectTaskCell Method (Project)

Selects a cell containing task information.


## Syntax

 _expression_. **SelectTaskCell**( ** _Row_**, ** _Column_**, ** _RowRelative_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Row_|Optional|**Long**|The row number (RowRelative is  **False** ) or the relative row position (RowRelative is **True** ) of the cell to select.|
| _Column_|Optional|**String**|The field name of the cell to select.|
| _RowRelative_|Optional|**Boolean**|**True** if the row number is relative to the active cell. The default value is **True**.|

### Return Value

 **Boolean**


## Remarks

Using the  **SelectTaskCell** method without specifying any arguments retains the current cell as the active cell. The **SelectTaskCell** method is only available when the Gantt Chart, Task Sheet, or Task Usage view is the active view.


