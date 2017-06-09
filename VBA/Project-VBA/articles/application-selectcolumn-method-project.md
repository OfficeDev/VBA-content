---
title: Application.SelectColumn Method (Project)
keywords: vbapj.chm2046
f1_keywords:
- vbapj.chm2046
ms.prod: project-server
api_name:
- Project.Application.SelectColumn
ms.assetid: 5bb674e9-253e-355f-a501-d0aeaef56535
ms.date: 06/08/2017
---


# Application.SelectColumn Method (Project)

Selects one or more columns.


## Syntax

 _expression_. **SelectColumn**( ** _Column_**, ** _Additional_**, ** _Extend_**, ** _Add_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Column_|Optional|**Integer**| The number of the column to select. (Columns are numbered from left to right, starting with 2.) The default is the active column.|
| _Additional_|Optional|**Integer**|The number of columns to select in addition to the active column.|
| _Extend_|Optional|**Boolean**|**True** if the active selection is extended into the new selection. The default value is **False**.|
| _Add_|Optional|**Boolean**|**True** if the new selection is added to the active selection. The default value is **False**.|

### Return Value

 **Boolean**


## Example

The following example selects the fourth column of the Gantt Chart.


```vb
Sub Select_Column() 
 ViewApply Name:="&;Gantt Chart" 
 SelectColumn Column:=4, Extend:=False 
End Sub
```


