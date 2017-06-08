---
title: ShadowFormat.IncrementOffsetX Method (Excel)
keywords: vbaxl10.chm114020
f1_keywords:
- vbaxl10.chm114020
ms.prod: excel
api_name:
- Excel.ShadowFormat.IncrementOffsetX
ms.assetid: eaa71500-16dd-5df1-cf32-920ab71d77bb
ms.date: 06/08/2017
---


# ShadowFormat.IncrementOffsetX Method (Excel)

Changes the horizontal offset of the shadow by the specified number of points. Use the  **[OffsetX](shadowformat-offsetx-property-excel.md)** property to set the absolute horizontal shadow offset.


## Syntax

 _expression_ . **IncrementOffsetX**( **_Increment_** )

 _expression_ A variable that represents a **ShadowFormat** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Increment_|Required| **Single**|Specifies how far the shadow offset is to be moved horizontally, in points. A positive value moves the shadow to the right; a negative value moves it to the left.|

## Example

This example moves the shadow on shape three on  `myDocument` to the left by 3 points.


```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes(3).Shadow.IncrementOffsetX -3
```


## See also


#### Concepts


[ShadowFormat Object](shadowformat-object-excel.md)

