---
title: Application.SelectCell Method (Project)
keywords: vbapj.chm2070
f1_keywords:
- vbapj.chm2070
ms.prod: project-server
api_name:
- Project.Application.SelectCell
ms.assetid: 7177d0bb-6e0e-8885-4f29-51faa34cea8b
ms.date: 06/08/2017
---


# Application.SelectCell Method (Project)

Selects a cell.


## Syntax

 _expression_. **SelectCell**( ** _Row_**, ** _Column_**, ** _RowRelative_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Row_|Optional|**Long**|The row number ( **RowRelative** is **False** ) or the relative row position ( **RowRelative** is **True** ) of the cell to select.|
| _Column_|Optional|**Integer**|The column number of the cell to select.|
| _RowRelative_|Optional|**Boolean**|**True** if the row number is relative to the active cell. The default value is **True**.|

### Return Value

 **Boolean**


## Remarks

Using the  **SelectCell** method without specifying any arguments retains the current cell as the active cell.


## Example

The following example selects the third field in the fourth row of the Gantt Chart.


```vb
Sub Select_Cell() 
 
 ViewApply Name:="&;Gantt Chart" 
 SelectCell Row:=4, Column:=3, RowRelative:=False 
End Sub
```


