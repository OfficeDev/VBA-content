---
title: Shape.Flip Method (Excel)
keywords: vbaxl10.chm636077
f1_keywords:
- vbaxl10.chm636077
ms.prod: excel
api_name:
- Excel.Shape.Flip
ms.assetid: 6ba41c89-878e-d9e1-5594-0cf45411b608
ms.date: 06/08/2017
---


# Shape.Flip Method (Excel)

Flips the specified shape around its horizontal or vertical axis.


## Syntax

 _expression_ . **Flip**( **_FlipCmd_** )

 _expression_ A variable that represents a **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FlipCmd_|Required| **[MsoFlipCmd](http://msdn.microsoft.com/library/8ca14f82-eaf6-754f-7a71-7b017dcfa230%28Office.15%29.aspx)**|Specifies whether the shape is to be flipped horizontally or vertically.|

## Example

This example adds a triangle to  `myDocument`, duplicates the triangle, and then flips the duplicate triangle vertically and makes it red.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes.AddShape(msoShapeRightTriangle, _ 
        10, 10, 50, 50).Duplicate 
    .Fill.ForeColor.RGB = RGB(255, 0, 0) 
    .Flip msoFlipVertical 
End With
```


## See also


#### Concepts


[Shape Object](shape-object-excel.md)

