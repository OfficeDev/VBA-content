---
title: ShapeRange.HasInkXML Property (PowerPoint)
ms.assetid: 1a7b7b8b-c0e8-8f62-1015-e99cb590fd50
ms.date: 06/08/2017
ms.prod: powerpoint
---


# ShapeRange.HasInkXML Property (PowerPoint)

Returns an [MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx) enumeration value that indicates whether the specified shape range contains ink XML that can be retrieved via the[ShapeRange.InkXML](shaperange-inkxml-property-powerpoint.md) property. Read-only.

An error is returned if the shape range does not contain any ink XML.

## Syntax

 _expression_. **HasInkXML**

 _expression_ A variable that represents a **ShapeRange** object.


### Return Value

MsoTriState


## Remarks

The value of the this property can be one of these  **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The specified shape range does not contain ink XML.|
|**msoTrue**| The specified shape range does not contain ink XML.|
|**msoTriStateMixed**| The specified shape range contains a mix of msoTrue and msoFalse return values. One or more shapes in the shape range contains ink XML and another shape in the shape range does not contain any ink XML.|

## See also


#### Concepts


[ShapeRange Object](shaperange-object-powerpoint.md)
#### Other resources


[MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx)
[ShapeRange.InkXML](shaperange-inkxml-property-powerpoint.md)


