---
title: ShadowFormat.IncrementOffsetY Method (Publisher)
keywords: vbapb10.chm3670033
f1_keywords:
- vbapb10.chm3670033
ms.prod: publisher
api_name:
- Publisher.ShadowFormat.IncrementOffsetY
ms.assetid: fca7a688-adf8-d8cd-8e14-9d1988c8d9f2
ms.date: 06/08/2017
---


# ShadowFormat.IncrementOffsetY Method (Publisher)

Incrementally changes the vertical offset of the shadow by the specified distance.


## Syntax

 _expression_. **IncrementOffsetY**( **_Increment_**)

 _expression_A variable that represents a  **ShadowFormat** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Increment|Required| **Variant**|Specifies how far the shadow offset is to be moved vertically. A positive value moves the shadow down; a negative value moves it up. Numeric values are evaluated in points; strings can be in any units supported by Microsoft Publisher (for example, "2.5 in").|

## Remarks

Use the  **[OffsetY](shadowformat-offsety-property-publisher.md)** property to set the absolute vertical shadow offset.

Use the  **[IncrementOffsetX](shadowformat-incrementoffsetx-method-publisher.md)** method to change a shadow's horizontal offset.


## Example

This example moves the shadow for the third shape in the active publication up by 3 points.


```vb
ActiveDocument.Pages(1).Shapes(3).Shadow _ 
 .IncrementOffsetY Increment:=-3 

```


