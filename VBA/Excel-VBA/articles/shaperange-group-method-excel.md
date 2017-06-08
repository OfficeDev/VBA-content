---
title: ShapeRange.Group Method (Excel)
keywords: vbaxl10.chm640086
f1_keywords:
- vbaxl10.chm640086
ms.prod: excel
api_name:
- Excel.ShapeRange.Group
ms.assetid: f0ad9b81-42ad-0ee6-d2e2-ff2a88d47a97
ms.date: 06/08/2017
---


# ShapeRange.Group Method (Excel)

Groups the shapes in the specified range.


## Syntax

 _expression_ . **Group**

 _expression_ A variable that represents a **ShapeRange** object.


### Return Value

A  **[Shape](shape-object-excel.md)** object that represents the grouped shape.


## Remarks

Because a group of shapes is treated as a single shape, grouping and ungrouping shapes changes the number of items in the  **[Shapes](shapes-object-excel.md)** collection and changes the index numbers of items that come after the affected items in the collection.

The  **[Range](range-object-excel.md)** object must be a single cell in the PivotTable field's data range. If you attempt to apply this method to more than one cell, it will fail (without displaying an error message).


## See also


#### Concepts


[ShapeRange Object](shaperange-object-excel.md)

