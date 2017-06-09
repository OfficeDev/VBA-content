---
title: Shape.LockAspectRatio Property (PowerPoint)
keywords: vbapp10.chm547028
f1_keywords:
- vbapp10.chm547028
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.LockAspectRatio
ms.assetid: b66acf40-1136-36b6-eabc-96b0fac527de
ms.date: 06/08/2017
---


# Shape.LockAspectRatio Property (PowerPoint)

Determines whether the specified shape retains its original proportions when you resize it. Read/write.


## Syntax

 _expression_. **LockAspectRatio**

 _expression_ A variable that represents a **Shape** object.


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

myDocument.Shapes.AddShape(msoShapeCube, 50, 50, 100, 200).LockAspectRatio = msoTrue
```


## See also


#### Concepts


[Shape Object](shape-object-powerpoint.md)

