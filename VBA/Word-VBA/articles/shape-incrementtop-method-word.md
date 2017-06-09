---
title: Shape.IncrementTop Method (Word)
keywords: vbawd10.chm161480720
f1_keywords:
- vbawd10.chm161480720
ms.prod: word
api_name:
- Word.Shape.IncrementTop
ms.assetid: 9aa5edb1-192f-5ccf-7513-3b9f660826ad
ms.date: 06/08/2017
---


# Shape.IncrementTop Method (Word)

Moves the specified shape vertically by the specified number of points.


## Syntax

 _expression_ . **IncrementTop**( **_Increment_** )

 _expression_ Required. A variable that represents a **[Shape](shape-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Increment_|Required| **Single**|Specifies how far the shape object is to be moved vertically, in points. A positive value moves the shape down; a negative value moves it up.|

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

