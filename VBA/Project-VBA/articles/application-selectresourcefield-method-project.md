---
title: Application.SelectResourceField Method (Project)
keywords: vbapj.chm2064
f1_keywords:
- vbapj.chm2064
ms.prod: project-server
api_name:
- Project.Application.SelectResourceField
ms.assetid: 6942d5a5-4072-4a95-f2b7-33bf965e302f
ms.date: 06/08/2017
---


# Application.SelectResourceField Method (Project)

Selects a resource field.


## Syntax

 _expression_. **SelectResourceField**( ** _Row_**, ** _Column_**, ** _RowRelative_**, ** _Width_**, ** _Height_**, ** _Extend_**, ** _Add_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Row_|Required|**Long**|The number of the row containing the field to select.|
| _Column_|Required|**String**|The name of the column containing the field to select.|
| _RowRelative_|Optional|**Boolean**|**True** if the location of the new selection is relative to the active selection. The default value is **True**.|
| _Width_|Optional|**Long**|The number of columns to select in addition to the active field.|
| _Height_|Optional|**Long**|The number of rows to select in addition to the active field.|
| _Extend_|Optional|**Boolean**|**True** if the active selection is extended into the new selection. The default value is **False**.|
| _Add_|Optional|**Boolean**|**True** if the new selection is added to the active selection. The default value is **False**.|

### Return Value

 **Boolean**


## Example

The following example selects the  **Name** column and the next two columns of the third and fourth rows.


```vb
Sub Select_ResourceField() 
 
 ViewApply Name:="&;Resource Sheet" 
 SelectResourceField Row:=3, Column:="Name", RowRelative:=False, Width:=2, Height:=1 
End Sub
```


