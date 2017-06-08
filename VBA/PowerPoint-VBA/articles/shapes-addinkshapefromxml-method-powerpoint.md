---
title: Shapes.AddInkShapeFromXML Method (PowerPoint)
ms.assetid: 88a395ac-b11e-d42e-f4b4-b41bf1d1347e
ms.date: 06/08/2017
ms.prod: powerpoint
---


# Shapes.AddInkShapeFromXML Method (PowerPoint)

Creates an ink shape. Returns a [Shape](shape-object-powerpoint.md) object that represents the new ink shape.


## Syntax

 _expression_. **AddInkShapeFromXML**( _InkXML_,  _InkXML_,  _Left_,  _Top_,  _Width_,  _Height_)

 _expression_ A variable that represents a **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _linkXML_|Required|**String**|The string that contains the InkActionML of the ink to create.|
| _Left_|Required|**Single**|The position, measured in points, of the left edge of the ink shape relative to the left edge of the slide.|
| _Top_|Required|**Single**|The position, measured in points, of the top edge of the ink shape relative to the top edge of the slide.|
| _Width_|Optional|**Single**| The width of the ink shape, measured in points. If this parameter is not specified, the width is calculated based off of the InkActionML.|
| _Height_|Optional|**Single**|The height of the ink shape, measured in points. If this parameter is not specified, the hight is calculated based off of the InkActionML.|

### Return Value

A [Shape](shape-object-powerpoint.md) object that represents the newly-added ink shape.


## See also


#### Concepts


[Shape](shape-object-powerpoint.md)
[Shapes Object](shapes-object-powerpoint.md)

