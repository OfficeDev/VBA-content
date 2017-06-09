---
title: ShapeRange.MergeShapes Method (Project)
ms.prod: project-server
ms.assetid: c470a800-6010-111b-831d-023e480fca31
ms.date: 06/08/2017
---


# ShapeRange.MergeShapes Method (Project)
The  **MergeShapes** method is not implemented in Project.

## Syntax

 _expression_. **MergeShapes** _(MergeCmd,_ _PrimaryShape)_

 _expression_ A variable that represents a **ShapeRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
| _MergeCmd_|Required|**[MsoMergeCmd](http://msdn.microsoft.com/en-us/library/office/jj227893%28v=office.15%29)**|The type of merge to perform.|
| _PrimaryShape_|Optional|**Shape**|The primary shape for the merge.|
| _MergeCmd_|Required|MSOMERGECMD||
| _PrimaryShape_|Optional|SHAPE||

### Return value

 **Nothing**


## Remarks

In general for applications that implement Office Art, the  **MergeShapes** method merges two or more shapes in a shape range into the specified **Shape** object. Project does not support the **MergeShapes** method.


## See also


#### Other resources


[ShapeRange Object](shaperange-object-project.md)
[MsoMergeCmd](http://msdn.microsoft.com/en-us/library/office/jj227893%28v=office.15%29)
