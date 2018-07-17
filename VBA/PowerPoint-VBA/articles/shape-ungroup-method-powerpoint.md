---
title: Shape.Ungroup Method (PowerPoint)
keywords: vbapp10.chm547013
f1_keywords:
- vbapp10.chm547013
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.Ungroup
ms.assetid: 2d0447df-7356-35e7-972e-e763ac1b3b8e
ms.date: 06/08/2017
---


# Shape.Ungroup Method (PowerPoint)

Ungroups any grouped shapes in the specified shape or range of shapes. Disassembles pictures and OLE objects within the specified shape or range of shapes. Returns the ungrouped shapes as a single  **[ShapeRange](shaperange-object-powerpoint.md)** object.


## Syntax

 _expression_. **Ungroup**

 _expression_ A variable that represents a **Shape** object.


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


[Shape Object](shape-object-powerpoint.md)

