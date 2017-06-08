---
title: Application.SelectRow Method (Project)
keywords: vbapj.chm2045
f1_keywords:
- vbapj.chm2045
ms.prod: project-server
api_name:
- Project.Application.SelectRow
ms.assetid: 63d31b23-3edb-9cd9-16c5-ac4ca4555a2c
ms.date: 06/08/2017
---


# Application.SelectRow Method (Project)

Selects one or more rows.


## Syntax

 _expression_. **SelectRow**( ** _Row_**, ** _RowRelative_**, ** _Height_**, ** _Extend_**, ** _Add_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Row_|Optional|**Long**|The number of the row to select. The default is the active row.|
| _RowRelative_|Optional|**Boolean**|**True** if the location of the new selection is relative to the active selection. The default value is **True**.|
| _Height_|Optional|**Long**|The number of rows to select in addition to the active cell.|
| _Extend_|Optional|**Boolean**|**True** if the active selection is extended into the new selection. The default value is **False**.|
| _Add_|Optional|**Boolean**|**True** if the new selection is added to the active selection. The default value is **False**.|

### Return Value

 **Boolean**


## Example

The following example adds row numbers 3 through 5 to the current selection.


```vb
Sub Select_Row() 
 
 'Activate Gantt Chart 
 ViewApply Name:="&;Gantt Chart" 
 SelectRow Row:=3, RowRelative:=False, Height:=2, Add:=True 
End Sub
```


