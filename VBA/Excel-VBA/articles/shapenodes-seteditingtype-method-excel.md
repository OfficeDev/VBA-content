---
title: ShapeNodes.SetEditingType Method (Excel)
keywords: vbaxl10.chm112009
f1_keywords:
- vbaxl10.chm112009
ms.prod: excel
api_name:
- Excel.ShapeNodes.SetEditingType
ms.assetid: 5bf464d6-b9d3-f62b-a625-0d153d7f265e
ms.date: 06/08/2017
---


# ShapeNodes.SetEditingType Method (Excel)

Sets the editing type of the node specified by  _Index_. If the node is a control point for a curved segment, this method sets the editing type of the node adjacent to it that joins two segments. Note that, depending on the editing type, this method may affect the position of adjacent nodes.


## Syntax

 _expression_ . **SetEditingType**( **_Index_** , **_EditingType_** )

 _expression_ A variable that represents a **ShapeNodes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|The node whose editing type is to be set.|
| _EditingType_|Required| **[MsoEditingType](http://msdn.microsoft.com/library/5fe5c4f6-6467-c6a7-197c-ff700c384b92%28Office.15%29.aspx)**|The editing property of the vertex.|

## Example

This example changes all corner nodes to smooth nodes in shape three on  `myDocument`. Shape three must be a freeform drawing.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(3).Nodes 
    For n = 1 to .Count 
        If .Item(n).EditingType = msoEditingCorner Then 
            .SetEditingType n, msoEditingSmooth 
        End If 
    Next 
End With
```


## See also


#### Concepts


[ShapeNodes Object](shapenodes-object-excel.md)

