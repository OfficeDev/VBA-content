---
title: ShapeRange.Regroup Method (PowerPoint)
keywords: vbapp10.chm548062
f1_keywords:
- vbapp10.chm548062
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange.Regroup
ms.assetid: 3da4a44d-4b0c-e335-b376-4d76fe5ed561
ms.date: 06/08/2017
---


# ShapeRange.Regroup Method (PowerPoint)

Regroups the group that the specified shape range belonged to previously. Returns the regrouped shapes as a single  **Shape** object.


## Syntax

 _expression_. **Regroup**

 _expression_ A variable that represents a **ShapeRange** object.


### Return Value

Shape


## Remarks

The  **Regroup** method only restores the group for the first previously grouped shape it finds in the specified **ShapeRange** collection. Therefore, if the specified shape range contains shapes that previously belonged to different groups, only one of the groups will be restored.

Note that because a group of shapes is treated as a single shape, grouping and ungrouping shapes changes the number of items in the  **Shapes** collection and changes the index numbers of items that come after the affected items in the collection.


## Example

This example regroups the shapes in the selection in the active window. If the shapes haven't been previously grouped and ungrouped, this example will fail.


```vb
ActiveWindow.Selection.ShapeRange.Regroup
```


## See also


#### Concepts


[ShapeRange Object](shaperange-object-powerpoint.md)

