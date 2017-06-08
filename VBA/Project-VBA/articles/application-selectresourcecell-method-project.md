---
title: Application.SelectResourceCell Method (Project)
keywords: vbapj.chm2069
f1_keywords:
- vbapj.chm2069
ms.prod: project-server
api_name:
- Project.Application.SelectResourceCell
ms.assetid: 3bae94f3-5661-63ef-47a6-12824d5426d0
ms.date: 06/08/2017
---


# Application.SelectResourceCell Method (Project)

Selects a cell containing resource information.


## Syntax

 _expression_. **SelectResourceCell**( ** _Row_**, ** _Column_**, ** _RowRelative_** )

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

Using the  **SelectResourceCell** method without specifying any arguments retains the current cell as the active cell. The **SelectResourceCell** method is only available when the Resource Sheet or Resource Usage view is the active view.


## Example

The following example selects the third row in the  **Name** column of the Resource Sheet.


```vb
Sub Select_ResourceCell() 
 
 ViewApply Name:="&;Resource Sheet" 
 SelectResourceCell Row:=3, Column:="Name", RowRelative:=False 
End Sub
```


