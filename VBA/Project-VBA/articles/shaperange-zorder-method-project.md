---
title: ShapeRange.ZOrder Method (Project)
ms.prod: project-server
ms.assetid: d713d882-a137-7fa2-0b2c-5b31f400eaa5
ms.date: 06/08/2017
---


# ShapeRange.ZOrder Method (Project)
Moves the shape range in front of or behind other shapes (that is, changes the position in the z-order).

## Syntax

 _expression_. **ZOrder** _(ZOrderCmd)_

 _expression_ A variable that represents a **ShapeRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ZOrderCmd_|Required|**[MsoZOrderCmd](http://msdn.microsoft.com/en-us/library/office/ff861432%28v=office.15%29)**|Specifies where to move the shape range relative to the other shapes.|
| _ZOrderCmd_|Required|MSOZORDERCMD||

### Return value

 **Nothing**


## Remarks

Use the  **ZOrderPosition** property to determine the current position of a shape in the z-order.


## See also


#### Other resources


[ShapeRange Object](shaperange-object-project.md)
[MsoZOrderCmd](http://msdn.microsoft.com/en-us/library/office/ff861432%28v=office.15%29)
[ZOrderPosition Property](shape-zorderposition-property-project.md)
