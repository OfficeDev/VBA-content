---
title: ShapeRange.IncrementTop Method (PowerPoint)
keywords: vbapp10.chm548007
f1_keywords:
- vbapp10.chm548007
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange.IncrementTop
ms.assetid: 55c18051-97a8-beab-c354-48256daff762
ms.date: 06/08/2017
---


# ShapeRange.IncrementTop Method (PowerPoint)

Moves the specified shape range vertically by the specified number of points.


## Syntax

 _expression_. **IncrementTop**( **_Increment_** )

 _expression_ A variable that represents a **ShapeRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Increment_|Required|**Single**|Specifies how far the shape range is to be moved vertically, in points. A positive value moves the shape range down; a negative value moves it up.|

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

