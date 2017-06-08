---
title: ShapeRange.Regroup Method (Excel)
keywords: vbaxl10.chm640089
f1_keywords:
- vbaxl10.chm640089
ms.prod: excel
api_name:
- Excel.ShapeRange.Regroup
ms.assetid: d30d3064-c37e-84b0-10a6-11dcd18c593e
ms.date: 06/08/2017
---


# ShapeRange.Regroup Method (Excel)

Regroups the group that the specified shape range belonged to previously. Returns the regrouped shapes as a single  **[Shape](shape-object-excel.md)** object.


## Syntax

 _expression_ . **Regroup**

 _expression_ A variable that represents a **ShapeRange** object.


### Return Value

Shape


## Remarks

The  **Regroup** method only restores the group for the first previously grouped shape it finds in the specified **[ShapeRange](shaperange-object-excel.md)** collection. Therefore, if the specified shape range contains shapes that previously belonged to different groups, only one of the groups will be restored.

Note that because a group of shapes is treated as a single shape, grouping and ungrouping shapes changes the number of items in the  **[Shapes](shapes-object-excel.md)** collection and changes the index numbers of items that come after the affected items in the collection.


## Example

This example regroups the shapes in the selection in the active window. If the shapes haven't been previously grouped and ungrouped, this example will fail.


```vb
ActiveWindow.Selection.ShapeRange.Regroup
```


## See also


#### Concepts


[ShapeRange Object](shaperange-object-excel.md)

