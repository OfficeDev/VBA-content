---
title: Databar.BarFillType Property (Excel)
keywords: vbaxl10.chm810091
f1_keywords:
- vbaxl10.chm810091
ms.prod: excel
api_name:
- Excel.Databar.BarFillType
ms.assetid: c83fc8d3-63aa-4989-8099-74bcad7d6fce
ms.date: 06/08/2017
---


# Databar.BarFillType Property (Excel)

Returns or sets how a data bar is filled with color. Read/write


## Syntax

 _expression_ . **BarFillType**

 _expression_ A variable that represents a **[Databar](databar-object-excel.md)** object.


### Return Value

 **[XlDataBarFillType](xldatabarfilltype-enumeration-excel.md)**


## Remarks

The default value of the  **BarFillType** property is **xlDataBarFillGradient** .


## Example

The following code example selects a range of cells, adds a data bar conditional formatting rule to that range, and then sets the data bar's fill color to solid.


```vb
Range("A1:A10").Select 
Range("A1:A10").Activate 
 
Set myDataBar = Selection.FormatConditions.AddDatabar 
myDataBar.BarFillType = xlDataBarFillSolid
```


## See also


#### Concepts


[Databar Object](databar-object-excel.md)

