---
title: Application.SelectTaskColumn Method (Project)
keywords: vbapj.chm2065
f1_keywords:
- vbapj.chm2065
ms.prod: project-server
api_name:
- Project.Application.SelectTaskColumn
ms.assetid: f4269749-de44-d7dd-de74-c95a046411fe
ms.date: 06/08/2017
---


# Application.SelectTaskColumn Method (Project)

Selects a column containing task information.


## Syntax

 _expression_. **SelectTaskColumn**( ** _Column_**, ** _Additional_**, ** _Extend_**, ** _Add_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Column_|Optional|**String**|The field name of the column to select. The default is the column containing the active cell.|
| _Additional_|Optional|**Integer**|The number of additional columns to select to the right of  **Column**. If **Extend** is **True**, **Additional** is ignored. The default value is 0.|
| _Extend_|Optional|**Boolean**|**True** if all columns between the current selection and **Column** are selected. The default value is **False**.|
| _Add_|Optional|**Boolean**|**True** if the current column is included in the selection. The default value is **False**.|

### Return Value

 **Boolean**


## Remarks

The  **SelectTaskColumn** method is only available when the Gantt Chart, Task Sheet, or Task Usage view is the active view.


