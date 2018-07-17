---
title: ShapeRange.IncrementLeft Method (Excel)
keywords: vbaxl10.chm640083
f1_keywords:
- vbaxl10.chm640083
ms.prod: excel
api_name:
- Excel.ShapeRange.IncrementLeft
ms.assetid: 604e8e92-b03a-da67-7022-4d73ebdf9872
ms.date: 06/08/2017
---


# ShapeRange.IncrementLeft Method (Excel)

Moves the specified shape horizontally by the specified number of points.


## Syntax

 _expression_ . **IncrementLeft**( **_Increment_** )

 _expression_ A variable that represents a **ShapeRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Increment_|Required| **Single**|Specifies how far the shape is to be moved horizontally, in points. A positive value moves the shape to the right; a negative value moves it to the left.|

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

