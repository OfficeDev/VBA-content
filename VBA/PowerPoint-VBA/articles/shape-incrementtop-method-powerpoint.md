---
title: Shape.IncrementTop Method (PowerPoint)
keywords: vbapp10.chm547007
f1_keywords:
- vbapp10.chm547007
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.IncrementTop
ms.assetid: b74307f9-9efa-4117-9232-24dfd2bdb883
ms.date: 06/08/2017
---


# Shape.IncrementTop Method (PowerPoint)

Moves the specified shape vertically by the specified number of points.


## Syntax

 _expression_. **IncrementTop**( **_Increment_** )

 _expression_ A variable that represents a **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Increment_|Required|**Single**|Specifies how far the shape is to be moved vertically, in points. A positive value moves the shape down; a negative value moves it up.|

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


[Shape Object](shape-object-powerpoint.md)

