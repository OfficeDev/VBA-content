---
title: ShapeRange.Ungroup Method (PowerPoint)
keywords: vbapp10.chm548013
f1_keywords:
- vbapp10.chm548013
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange.Ungroup
ms.assetid: 7bac0e8b-09d5-b219-af20-2a3b8dcee9d9
ms.date: 06/08/2017
---


# ShapeRange.Ungroup Method (PowerPoint)

Ungroups any grouped shapes in the specified shape or range of shapes. Disassembles pictures and OLE objects within the specified shape or range of shapes. Returns the ungrouped shapes as a single  **[ShapeRange](shaperange-object-powerpoint.md)** object.


## Syntax

 _expression_. **Ungroup**

 _expression_ A variable that represents a **ShapeRange** object.


### Return Value

ShapeRange


## Remarks

Because a group of shapes is treated as a single object, grouping and ungrouping shapes changes the number of items in the  **Shapes** collection and changes the index numbers of items that come after the affected items in the collection.


## Example

This example ungroups any grouped shapes and disassembles any pictures or OLE objects on  `myDocument`.


```vb
Set myDocument = ActivePresentation.Slides(1)

For Each s In myDocument.Shapes

    s.Ungroup

Next
```

This example ungroups any grouped shapes on  `myDocument` without disassembling pictures or OLE objects on the slide.




```vb
Set myDocument = ActivePresentation.Slides(1)

For Each s In myDocument.Shapes

    If s.Type = msoGroup Then s.Ungroup

Next
```


## See also


#### Concepts


[ShapeRange Object](shaperange-object-powerpoint.md)

