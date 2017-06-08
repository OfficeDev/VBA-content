---
title: Shape.IncrementLeft Method (Word)
keywords: vbawd10.chm161480718
f1_keywords:
- vbawd10.chm161480718
ms.prod: word
api_name:
- Word.Shape.IncrementLeft
ms.assetid: e3073ce8-7d72-1520-e042-c4b392fae460
ms.date: 06/08/2017
---


# Shape.IncrementLeft Method (Word)

Moves the specified shape horizontally by the specified number of points.


## Syntax

 _expression_ . **IncrementLeft**( **_Increment_** )

 _expression_ Required. A variable that represents a **[Shape](shape-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Increment_|Required| **Single**|Specifies how far the shape is to be moved horizontally, in points. A positive value moves the shape to the right; a negative value moves it to the left.|

## Example

This example duplicates shape one on  _myDocument_ , sets the fill for the duplicate, moves it 70 points to the right and 50 points up, and rotates it 30 degrees clockwise.


```vb
Set myDocument = ActiveDocument 
With myDocument.Shapes(1).Duplicate 
 .Fill.PresetTextured msoTextureGranite 
 .IncrementLeft 70 
 .IncrementTop -50 
 .IncrementRotation 30 
End With
```


## See also


#### Concepts


[Shape Object](shape-object-word.md)

