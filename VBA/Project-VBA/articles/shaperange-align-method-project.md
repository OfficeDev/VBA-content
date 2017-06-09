---
title: ShapeRange.Align Method (Project)
ms.prod: project-server
ms.assetid: 6e8e3a02-efd8-995c-be1a-a89d7709bd08
ms.date: 06/08/2017
---


# ShapeRange.Align Method (Project)
The  **Align** method is not implemented in Project.

## Syntax

 _expression_. **Align** _(AlignCmd,_ _RelativeTo)_

 _expression_ A variable that represents a **ShapeRange** object.


### Return value

 **Nothing**


## Remarks

In general for applications that implement Office Art, the  **Align** method aligns the shapes contained in the shape range. Project does not support automatic distribution or alignment of shapes in a report.

If you try to use the  **Align** method, such as `sRange1.Align msoAlignMiddles, msoFalse`, you get run-time error &;H80070057, "The specified value is out of range."


## See also


#### Other resources


[ShapeRange Object](shaperange-object-project.md)
