---
title: ShadowFormat.IncrementOffsetY Method (PowerPoint)
keywords: vbapp10.chm554003
f1_keywords:
- vbapp10.chm554003
ms.prod: powerpoint
api_name:
- PowerPoint.ShadowFormat.IncrementOffsetY
ms.assetid: a220a04d-90d1-1788-b4d9-5b9af5739c69
ms.date: 06/08/2017
---


# ShadowFormat.IncrementOffsetY Method (PowerPoint)

Changes the vertical offset of the shadow by the specified number of points. 


## Syntax

 _expression_. **IncrementOffsetY**( **_Increment_** )

 _expression_ A variable that represents an **ShadowFormat** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Increment_|Required|**Single**|Specifies how far the shadow offset is to be moved vertically, in points. A positive value moves the shadow down; a negative value moves it up.|

## Remarks

Use the  **[OffsetY](shadowformat-offsety-property-powerpoint.md)** property to set the absolute vertical shadow offset.


## Example

This example moves the shadow for shape three on  `myDocument` up by 3 points.


```vb
Set myDocument = ActivePresentation.Slides(1)

myDocument.Shapes(3).Shadow.IncrementOffsetY -3
```


## See also


#### Concepts


[ShadowFormat Object](shadowformat-object-powerpoint.md)

