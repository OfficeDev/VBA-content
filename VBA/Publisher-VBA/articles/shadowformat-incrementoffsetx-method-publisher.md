---
title: ShadowFormat.IncrementOffsetX Method (Publisher)
keywords: vbapb10.chm3670032
f1_keywords:
- vbapb10.chm3670032
ms.prod: publisher
api_name:
- Publisher.ShadowFormat.IncrementOffsetX
ms.assetid: 05c25f0f-beac-2b25-630b-57d4a3bdb0c9
ms.date: 06/08/2017
---


# ShadowFormat.IncrementOffsetX Method (Publisher)

Incrementally changes the horizontal offset of the shadow by the specified distance.


## Syntax

 _expression_. **IncrementOffsetX**( **_Increment_**)

 _expression_A variable that represents a  **ShadowFormat** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Increment|Required| **Variant**|Specifies how far the shadow offset is to be moved horizontally. A positive value moves the shadow to the right; a negative value moves it to the left. Numeric values are evaluated in points; strings can be in any units supported by Microsoft Publisher (for example, "2.5 in").|

## Remarks

Use the  **[OffsetX](shadowformat-offsetx-property-publisher.md)** property to set the absolute horizontal shadow offset.

Use the  **[IncrementOffsetY](shadowformat-incrementoffsety-method-publisher.md)** method to change a shadow's vertical offset.


## Example

This example moves the shadow for the third shape in the active publication to the left by 3 points.


```vb
ActiveDocument.Pages(1).Shapes(3).Shadow _ 
 .IncrementOffsetX Increment:=-3 

```


