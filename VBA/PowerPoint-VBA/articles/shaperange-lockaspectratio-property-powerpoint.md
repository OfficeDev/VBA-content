---
title: ShapeRange.LockAspectRatio Property (PowerPoint)
keywords: vbapp10.chm548028
f1_keywords:
- vbapp10.chm548028
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange.LockAspectRatio
ms.assetid: e30f2834-b6c2-d966-dbee-b22912e4e3f0
ms.date: 06/08/2017
---


# ShapeRange.LockAspectRatio Property (PowerPoint)

Determines whether the specified shape retains its original proportions when you resize it. Read/write.


## Syntax

 _expression_. **LockAspectRatio**

 _expression_ A variable that represents a **ShapeRange** object.


### Return Value

MsoTriState


## Remarks

The value of the  **LockAspectRatio** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|You can change the height and width of the shape independently of one another when you resize it.|
|**msoTrue**| The specified shape retains its original proportions when you resize it.|

## Example

This example adds a cube to  `myDocument`. The cube can be moved and resized, but not reproportioned.


```vb
Set myDocument = ActivePresentation.Slides(1)

myDocument.Shapes.AddShape(msoShapeCube, 50, 50, 100, 200) _
    .LockAspectRatio = msoTrue
```


## See also


#### Concepts


[ShapeRange Object](shaperange-object-powerpoint.md)

