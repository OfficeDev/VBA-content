---
title: ShapeRange.IncrementLeft Method (Publisher)
keywords: vbapb10.chm2293792
f1_keywords:
- vbapb10.chm2293792
ms.prod: publisher
api_name:
- Publisher.ShapeRange.IncrementLeft
ms.assetid: 1b760b5d-9879-5f64-c4c5-c9834a7928ff
ms.date: 06/08/2017
---


# ShapeRange.IncrementLeft Method (Publisher)

Moves the specified shape or shape range horizontally by the specified distance.


## Syntax

 _expression_. **IncrementLeft**( **_Increment_**)

 _expression_A variable that represents a  **ShapeRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Increment|Required| **Variant**|The horizontal distance to move the shape or shape range. A positive value moves the shape or shape range to the right; a negative value moves it to the left. Numeric values are evaluated in points; strings can be in any units supported by Microsoft Publisher (for example, "2.5 in").|

### Return Value

Nothing


## Remarks

Use the  **[IncrementTop](shape-incrementtop-method-publisher.md)** method to move shapes or shape ranges vertically.


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


