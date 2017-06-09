---
title: ShapeRange.Duplicate Method (PowerPoint)
keywords: vbapp10.chm548053
f1_keywords:
- vbapp10.chm548053
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange.Duplicate
ms.assetid: da7e1e45-480d-313d-1d12-65ee5bf26d86
ms.date: 06/08/2017
---


# ShapeRange.Duplicate Method (PowerPoint)

Creates a duplicate of the specified  **ShapeRange** object, adds the range of shapes to the **Shapes** collection, and then returns the new **ShapeRange** object. The duplicated objects are placed at the end of the **Shapes** collection.


## Syntax

 _expression_. **Duplicate**

 _expression_ A variable that represents a **ShapeRange** object.


### Return Value

ShapeRange


## Example

This example adds a new, blank slide at the end of the active presentation, adds a diamond shape to the new slide, duplicates the diamond, and then sets properties for the duplicate. The first diamond will have the default fill color for the active color scheme; the second diamond will be offset from the first one and will have the default shadow color.


```vb
Set mySlides = ActivePresentation.Slides
Set newSlide = mySlides.Add(mySlides.Count + 1, ppLayoutBlank)
Set firstObj = newSlide.Shapes _
    .AddShape(msoShapeDiamond, 10, 10, 250, 350)

With firstObj.Duplicate
    .Left = 150
    .Fill.ForeColor.SchemeColor = ppShadow
End With
```


## See also


#### Concepts


[ShapeRange Object](shaperange-object-powerpoint.md)

