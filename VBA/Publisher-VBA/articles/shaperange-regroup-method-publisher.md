---
title: ShapeRange.Regroup Method (Publisher)
keywords: vbapb10.chm2294019
f1_keywords:
- vbapb10.chm2294019
ms.prod: publisher
api_name:
- Publisher.ShapeRange.Regroup
ms.assetid: 29342a78-9425-2356-963c-36a62a7f3091
ms.date: 06/08/2017
---


# ShapeRange.Regroup Method (Publisher)

Regroups the group that the specified shape range belonged to previously. Returns the regrouped shapes as a single  **Shape** object.


## Syntax

 _expression_. **Regroup**

 _expression_A variable that represents a  **ShapeRange** object.


### Return Value

Shape


## Remarks

The  **Regroup** method only restores the group for the first previously grouped shape it finds in the specified **ShapeRange** collection. Therefore, if the specified shape range contains shapes that previously belonged to different groups, only one of the groups will be restored.

An error occurs if none of the shapes in the specified range were previously members of a group.

Because a group of shapes is treated as a single shape, grouping and ungrouping shapes changes the number of items in the  **Shapes** collection and changes the index numbers of items that come after the affected items in the collection.


## Example

This example regroups the selected shapes in the active publication. If the shapes haven't been previously grouped and ungrouped, this example will fail.


```vb
ActiveDocument.Selection.ShapeRange.Regroup
```


