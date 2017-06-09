---
title: Shape.HasInkXML Property (PowerPoint)
ms.assetid: 3d985f9b-64e3-8712-fd5f-73d38ca56810
ms.date: 06/08/2017
ms.prod: powerpoint
---


# Shape.HasInkXML Property (PowerPoint)

Returns an [MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx) enumeration value that indicates whether the specified shape contains ink XML that can be retrieved via the[Shape.InkXML](shape-inkxml-property-powerpoint.md) property. Read-only.

An error is returned if the shape does not contain any ink XML.

## Syntax

 _expression_. **HasInkXML**

 _expression_ A variable that represents a **Shape** object.


### Return Value

MsoTriState


## Remarks

The value of this property can be one of these  **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The specified shape does not contain ink XML.|
|**msoTrue**| The specified shape contains ink XML.|

## See also


#### Concepts


[Shape Object](shape-object-powerpoint.md)
#### Other resources


[MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx)
[Shape.InkXML](shape-inkxml-property-powerpoint.md)


