---
title: ShapeRange.Distribute Method (Project)
ms.prod: project-server
ms.assetid: 149081d5-8826-1395-e838-1333a4233981
ms.date: 06/08/2017
---


# ShapeRange.Distribute Method (Project)
The  **Distribute** method is not implemented in Project.

## Syntax

 _expression_. **Distribute** _(DistributeCmd,_ _RelativeTo)_

 _expression_ A variable that represents a **ShapeRange** object.


### Return value

 **Nothing**


## Remarks

In general for applications that implement Office Art,, the  **Distribute** method evenly distributes the shapes contained in the shape range. Project does not support automatic distribution or alignment of shapes in a report.

If you try to use the  **Distribute** method, such as `sRange1.Distribute msoDistributeHorizontally, msoFalse`, you get run-time error &;H80070057, "The specified value is out of range."


## See also


#### Other resources


[ShapeRange Object](shaperange-object-project.md)
