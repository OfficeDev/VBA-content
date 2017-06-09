---
title: Shape.IncrementTop Method (Publisher)
keywords: vbapb10.chm2228258
f1_keywords:
- vbapb10.chm2228258
ms.prod: publisher
api_name:
- Publisher.Shape.IncrementTop
ms.assetid: c7a5bf47-7c5a-f6e8-b2b7-c95bea9dc081
ms.date: 06/08/2017
---


# Shape.IncrementTop Method (Publisher)

Moves the specified shape or shape range vertically by the specified distance.


## Syntax

 _expression_. **IncrementTop**( **_Increment_**)

 _expression_A variable that represents a  **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Increment|Required| **Variant**|The vertical distance to move the shape or shape range. A positive value moves the shape or shape range down; a negative value moves it up. Numeric values are evaluated in points; strings can be in any units supported by Microsoft Publisher (for example, "2.5 in").|

## Remarks

Use the  **[IncrementLeft](shape-incrementleft-method-publisher.md)** method to move shapes or shape ranges horizontally.


## Example

This example duplicates the first shape on the active publication, sets the fill for the duplicate, moves it 70 points to the right and 50 points up, and rotates it 30 degrees clockwise.


```vb
With ActiveDocument.Pages(1).Shapes(1).Duplicate 
 .Fill.PresetTextured PresetTexture:=msoTextureGranite 
 .IncrementLeft Increment:=70 
 .IncrementTop Increment:=-50 
 .IncrementRotation Increment:=30 
End With
```


