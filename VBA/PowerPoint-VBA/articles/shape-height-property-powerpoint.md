---
title: Shape.Height Property (PowerPoint)
keywords: vbapp10.chm547024
f1_keywords:
- vbapp10.chm547024
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.Height
ms.assetid: 3e2e7adf-9115-a903-c119-6429a10cbd9e
ms.date: 06/08/2017
---


# Shape.Height Property (PowerPoint)

Returns or sets the height of the specified object, in points. Read/write.


## Syntax

 _expression_. **Height**

 _expression_ A variable that represents a **Shape** object.


### Return Value

Single


## Remarks

The  **Height** property of a **Shape** object returns or sets the height of the forward-facing surface of the specified shape. This measurement doesn't include shadows or 3-D effects.


## Example

This example sets the height of document window two to half the height of the application window.


```
Windows(2).Height = Application.Height / 2
```

This example sets the height for row two in the specified table to 100 points (72 points per inch).




```vb
ActivePresentation.Slides(2).Shapes(5).Table.Rows(2).Height = 100
```


## See also


#### Concepts


[Shape Object](shape-object-powerpoint.md)

