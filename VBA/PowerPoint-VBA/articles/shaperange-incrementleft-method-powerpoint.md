---
title: ShapeRange.IncrementLeft Method (PowerPoint)
keywords: vbapp10.chm548005
f1_keywords:
- vbapp10.chm548005
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange.IncrementLeft
ms.assetid: 08d84101-bdfe-c3c6-a309-00c2fb2adab5
ms.date: 06/08/2017
---


# ShapeRange.IncrementLeft Method (PowerPoint)

Moves the specified shape range horizontally by the specified number of points.


## Syntax

 _expression_. **IncrementLeft**( **_Increment_** )

 _expression_ A variable that represents a **ShapeRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Increment_|Required|**Single**|Specifies how far the shape range is to be moved horizontally, in points. A positive value moves the shape range to the right; a negative value moves it to the left.|

## Example

This example duplicates shape one on  `myDocument`, sets the fill for the duplicate, moves it 70 points to the right and 50 points up, and rotates it 30 degrees clockwise.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(1).Duplicate

    .Fill.PresetTextured msoTextureGranite

    .IncrementLeft 70

    .IncrementTop -50

    .IncrementRotation 30

End With
```


## See also


#### Concepts


[ShapeRange Object](shaperange-object-powerpoint.md)

