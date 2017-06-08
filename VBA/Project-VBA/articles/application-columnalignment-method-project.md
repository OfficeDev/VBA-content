---
title: Application.ColumnAlignment Method (Project)
keywords: vbapj.chm2325
f1_keywords:
- vbapj.chm2325
ms.prod: project-server
api_name:
- Project.Application.ColumnAlignment
ms.assetid: 9c51eb2d-c28b-cb00-57e5-1643093e4acb
ms.date: 06/08/2017
---


# Application.ColumnAlignment Method (Project)

Sets the alignment of text in the active columns.


## Syntax

 _expression_. **ColumnAlignment**( ** _Align_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Align_|Required|**Long**|The alignment of text in the active columns. Can be one of the following  **PjAlignment** constants: **pjLeft**, **pjCenter**, or **pjRight**.|

### Return Value

 **Boolean**


## Example

The following example aligns the Start column to the left side.


```vb
Sub Column_Alignment() 
 
 'Activate Gantt Chart view 
 ViewApply Name:="Gantt Chart" 
 
 SelectTaskColumn Column:="Start" 
 ColumnAlignment Align:=pjLeft 
End Sub
```


