---
title: ShadowFormat.IncrementOffsetX Method (PowerPoint)
keywords: vbapp10.chm554002
f1_keywords:
- vbapp10.chm554002
ms.prod: powerpoint
api_name:
- PowerPoint.ShadowFormat.IncrementOffsetX
ms.assetid: 29fbda10-d3ed-963f-364d-5a5bbce92f34
ms.date: 06/08/2017
---


# ShadowFormat.IncrementOffsetX Method (PowerPoint)

Changes the horizontal offset of the shadow by the specified number of points. 


## Syntax

 _expression_. **IncrementOffsetX**( **_Increment_** )

 _expression_ A variable that represents an **ShadowFormat** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Increment_|Required|**Single**|Specifies how far the shadow offset is to be moved horizontally, in points. A positive value moves the shadow to the right; a negative value moves it to the left.|

## Remarks

Use the  **[OffsetX](shadowformat-offsetx-property-powerpoint.md)** property to set the absolute horizontal shadow offset.


## Example

This example moves the shadow for shape three on  `myDocument` to the left by 3 points.


```vb
Set myDocument = ActivePresentation.Slides(1)

myDocument.Shapes(3).Shadow.IncrementOffsetX -3
```


## See also


#### Concepts


[ShadowFormat Object](shadowformat-object-powerpoint.md)

