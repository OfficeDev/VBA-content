---
title: Application.SelectBeginning Method (Project)
keywords: vbapj.chm2041
f1_keywords:
- vbapj.chm2041
ms.prod: project-server
api_name:
- Project.Application.SelectBeginning
ms.assetid: 4adf20ae-4fd2-818a-da8c-133c08cad7fb
ms.date: 06/08/2017
---


# Application.SelectBeginning Method (Project)

Selects the first cell in the active table or view.


## Syntax

 _expression_. **SelectBeginning**( ** _Extend_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Extend_|Optional|**Boolean**|**True** if the current selection is extended to the first cell. If the active view is the Network Diagram or Resource Graph, Extend is ignored. The default value is **False**.|

### Return Value

 **Boolean**


## Remarks

In the Resource Graph,  **SelectBeginning** selects the resource with the lowest identification number. In the Network Diagram, **SelectBeginning** selects the box closest to the upper-left corner of the view.


## Example

The following example selects the "Name" field of row 4 as the beginning field in the Gantt Chart.


```vb
Sub Select_Beginning() 
 
 ViewApply Name:="&;Gantt Chart" 
 SelectTaskField Row:=4, Column:="Name", RowRelative:=False 
 
 SelectBeginning Extend:=True 
End Sub
```


