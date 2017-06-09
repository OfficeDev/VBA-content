---
title: Shape.ZOrder Method (Project)
ms.prod: project-server
ms.assetid: e8badff9-fbe5-b6b8-8c33-68cfde3bef38
ms.date: 06/08/2017
---


# Shape.ZOrder Method (Project)
Moves the shape in front of or behind other shapes (that is, changes the position in the z-order).

## Syntax

 _expression_. **ZOrder** _(ZOrderCmd)_

 _expression_ A variable that represents a **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ZOrderCmd_|Required|**[MsoZOrderCmd](http://msdn.microsoft.com/en-us/library/office/ff861432%28v=office.15%29)**|Specifies where to move the shape relative to the other shapes.|
| _ZOrderCmd_|Required|MSOZORDERCMD||

### Return value

 **Nothing**


## Remarks

Use the  **ZOrderPosition** property to determine the current position of a shape in the z-order.


## See also


#### Other resources


[Shape Object](shape-object-project.md)
[MsoZOrderCmd](http://msdn.microsoft.com/en-us/library/office/ff861432%28v=office.15%29)
[ZOrderPosition Property](shaperange-zorderposition-property-project.md)
