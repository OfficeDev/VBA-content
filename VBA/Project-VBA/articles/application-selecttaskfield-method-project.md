---
title: Application.SelectTaskField Method (Project)
keywords: vbapj.chm2063
f1_keywords:
- vbapj.chm2063
ms.prod: project-server
api_name:
- Project.Application.SelectTaskField
ms.assetid: 182bfb43-c1ae-32e1-2e93-7cb035e36bd0
ms.date: 06/08/2017
---


# Application.SelectTaskField Method (Project)

Selects a task field.


## Syntax

 _expression_. **SelectTaskField**( ** _Row_**, ** _Column_**, ** _RowRelative_**, ** _Width_**, ** _Height_**, ** _Extend_**, ** _Add_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Row_|Required|**Long**|The number of the row containing the field to select.|
| _Column_|Required|**String**|The name of the column containing the field to select.|
| _RowRelative_|Optional|**Boolean**|**True** if the location of the new selection is relative to the active selection. The default value is **True**.|
| _Width_|Optional|**Long**|The number of columns to select in addition to the active field.|
| _Height_|Optional|**Long**|The number of rows to select in addition to the active field.|
| _Extend_|Optional|**Boolean**|**True** if the active selection is extended into the new selection. The default value is **False.**|
| _Add_|Optional|**Boolean**|**True** if the new selection is added to the active selection. The default value is **False**.|

### Return Value

 **Boolean**


## Example

The following example selects the  **Name** column and the next two columns of the third and fourth rows on the Gantt Chart.


```vb
Sub Select_TaskField() 
 
 ViewApply Name:="&;Gantt Chart" 
 SelectTaskField Row:=3, Column:="Name", RowRelative:=False, Width:=2, Height:=1 
End Sub
```


