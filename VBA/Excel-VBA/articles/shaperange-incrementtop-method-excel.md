---
title: ShapeRange.IncrementTop Method (Excel)
keywords: vbaxl10.chm640085
f1_keywords:
- vbaxl10.chm640085
ms.prod: excel
api_name:
- Excel.ShapeRange.IncrementTop
ms.assetid: 39004de1-dbae-b57b-e2ea-edfc9b3aa9e3
ms.date: 06/08/2017
---


# ShapeRange.IncrementTop Method (Excel)

Moves the specified shape vertically by the specified number of points.


## Syntax

 _expression_ . **IncrementTop**( **_Increment_** )

 _expression_ A variable that represents a **ShapeRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Increment_|Required| **Single**|Specifies how far the shape object is to be moved vertically, in points. A positive value moves the shape down; a negative value moves it up.|

## Example

This example duplicates shape one on  `myDocument`, sets the fill for the duplicate, moves it 70 points to the right and 50 points up, and rotates it 30 degrees clockwise.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(1).Duplicate 
 .Fill.PresetTextured msoTextureGranite 
 .IncrementLeft 70 
 .IncrementTop -50 
 .IncrementRotation 30 
End With
```


## See also


#### Concepts


[ShapeRange Object](shaperange-object-excel.md)

