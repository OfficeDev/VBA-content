---
title: Shape.Duplicate Method (PowerPoint)
keywords: vbapp10.chm547053
f1_keywords:
- vbapp10.chm547053
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.Duplicate
ms.assetid: 0d2f22bc-ee72-6405-011a-77a9eb98fb39
ms.date: 06/08/2017
---


# Shape.Duplicate Method (PowerPoint)

Creates a duplicate of the specified  **Shape** object, adds the new shape to the **Shapes** collection, and then returns a new **ShapeRange** object. The duplicated objects are placed at the end of the **Shapes** collection.


## Syntax

 _expression_. **Duplicate**

 _expression_ A variable that represents a **Shape** object.


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


[Shape Object](shape-object-powerpoint.md)

