---
title: Application.SelectRange Method (Project)
keywords: vbapj.chm2062
f1_keywords:
- vbapj.chm2062
ms.prod: project-server
api_name:
- Project.Application.SelectRange
ms.assetid: 16b5925e-393b-3d4f-70d4-89213f521485
ms.date: 06/08/2017
---


# Application.SelectRange Method (Project)

Selects one or more cells.


## Syntax

 _expression_. **SelectRange**( ** _Row_**, ** _Column_**, ** _RowRelative_**, ** _Width_**, ** _Height_**, ** _Extend_**, ** _Add_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Row_|Required|**Long**|The number of the row containing the cell to select.|
| _Column_|Required|**Integer**|The number of the column containing the cell to select. (Columns are numbered from left to right, starting with 2.)|
| _RowRelative_|Optional|**Boolean**|**True** if the location of the new selection is relative to the active selection. The default value is **True**.|
| _Width_|Optional|**Long**|The number of columns to select in addition to the active cell.|
| _Height_|Optional|**Long**|The number of rows to select in addition to the active cell.|
| _Extend_|Optional|**Boolean**|**True** if the active selection is extended into the new selection. The default value is **False**.|
| _Add_|Optional|**Boolean**|**True** if the new selection is added to the active selection. The default value is **False**.|

### Return Value

 **Boolean**


## Example

The following example selects columns 3 through 6 and rows 4 through 6 on the Gantt Chart.


```vb
Sub Select_Range() 
 
 ViewApply Name:="&;Gantt Chart" 
 SelectRange Row:=4, Column:=3, RowRelative:=False, Width:=3, Height:=2 
 
End Sub
```


