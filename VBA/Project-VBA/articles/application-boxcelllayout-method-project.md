---
title: Application.BoxCellLayout Method (Project)
keywords: vbapj.chm2392
f1_keywords:
- vbapj.chm2392
ms.prod: project-server
api_name:
- Project.Application.BoxCellLayout
ms.assetid: 9b1ab0f5-d3ef-3258-aa01-ae1dea264ec5
ms.date: 06/08/2017
---


# Application.BoxCellLayout Method (Project)

Sets the cell layout and size properties for a data template in the Network Diagram view. The initial layout of a new data template is 2 rows by 2 columns of 100% width cells with cell merging enabled.


## Syntax

 _expression_. **BoxCellLayout**( ** _Name_**, ** _CellRows_**, ** _CellColumns_**, ** _CellWidth_**, ** _MergeCells_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required|**String**|**String**. The name of the data template to edit.|
| _CellRows_|Optional|**Long**|A value from 1 through 4 that specifies the number of rows of cells in the data template.|
| _CellColumns_|Optional|**Long**|A value from 1 through 4 that specifies the number of columns of cells in the data template.|
| _CellWidth_|Optional|**Integer**|A value from 100 through 200 that specifies the percentage by which to enlarge the width of the template cells.|
| _MergeCells_|Optional|**Boolean**|**True** if blank cells are merged with the cell to the left.|

### Return Value

 **Boolean**


## Remarks

Using the  **BoxCellLayout** method with only the _Name_ argument specified has no effect.


## Example

The following example modifies a copy of the  **Critical** data template named **Test Critical**. The macro removes the fourth row of cells and sets the fourth cell in the third row to show the  **Actual Cost** field and label in a purple-blue color.


```vb
Sub ModifyCriticalDataTemplate() 
 Application.BoxCellLayout Name:="Test Critical", CellRows:=3, MergeCells:=True 
 
 Application.BoxCellEditEx Name:="Test Critical", Cell:=pjCell4_3, _ 
 FieldName:=PjField.pjTaskActualCost, Font:="Arial", FontSize:="8", FontColor:=&;HFF0077, _ 
 Bold:=False, Italic:=False, Underline:=False, HorizontalAlignment:=pjLeft, _ 
 VerticalAlignment:=pjMiddle, TextLineLimit:=1, ShowLabel:=True, Label:="Cost" 
End Sub
```


