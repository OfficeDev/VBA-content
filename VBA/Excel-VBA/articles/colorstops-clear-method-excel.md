---
title: ColorStops.Clear Method (Excel)
keywords: vbaxl10.chm853078
f1_keywords:
- vbaxl10.chm853078
ms.prod: excel
api_name:
- Excel.ColorStops.Clear
ms.assetid: 308edcb7-6085-77d6-5e6a-d8ec1d31c043
ms.date: 06/08/2017
---


# ColorStops.Clear Method (Excel)

Clears the represented object.


## Syntax

 _expression_ . **Clear**

 _expression_ An expression that returns a **ColorStops** object.


### Return Value

Nothing


## Example

Clears the current ColorStops


```vb
Range("A1:A10").Select 
With Selection.Interior 
 .Pattern = xlPatternLinearGradient 
 .Gradient.Degree = 90 
 .Gradient.ColorStops.Clear 
End With
```


## See also


#### Concepts


[ColorStops Object](colorstops-object-excel.md)

