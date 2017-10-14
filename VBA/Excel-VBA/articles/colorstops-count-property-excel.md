---
title: ColorStops.Count Property (Excel)
keywords: vbaxl10.chm853073
f1_keywords:
- vbaxl10.chm853073
ms.prod: excel
api_name:
- Excel.ColorStops.Count
ms.assetid: 0574a698-ff87-56e3-eea9-aa2e6e77f270
ms.date: 06/08/2017
---


# ColorStops.Count Property (Excel)

Returns or sets the count of the represented object. Read-only


## Syntax

 _expression_ . **Count**

 _expression_ An expression that returns a **ColorStops** object.


### Return Value

Long


## Example

Returns the number of ColorStops in the active cell.


```vb
ActiveCell.Interior.Gradient.ColorStops.Count
```


## See also


#### Concepts


[ColorStops Object](colorstops-object-excel.md)

