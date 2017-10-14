---
title: Shape.Table Property (PowerPoint)
keywords: vbapp10.chm547060
f1_keywords:
- vbapp10.chm547060
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.Table
ms.assetid: cc57c50b-8c88-d863-31d2-a758eff5359f
ms.date: 06/08/2017
---


# Shape.Table Property (PowerPoint)

Returns a  **[Table](table-object-powerpoint.md)** object that represents a table in a shape or in a shape range. Read-only.


## Syntax

 _expression_. **Table**

 _expression_ A variable that represents a **Shape** object.


### Return Value

Table


## Example

This example sets the width of the first column in the table in shape five on the second slide to 80 points.


```vb
ActivePresentation.Slides(2).Shapes(5).Table.Columns(1).Width = 80
```


## See also


#### Concepts


[Shape Object](shape-object-powerpoint.md)

