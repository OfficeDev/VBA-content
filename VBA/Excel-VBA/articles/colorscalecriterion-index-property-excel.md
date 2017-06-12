---
title: ColorScaleCriterion.Index Property (Excel)
keywords: vbaxl10.chm808073
f1_keywords:
- vbaxl10.chm808073
ms.prod: excel
api_name:
- Excel.ColorScaleCriterion.Index
ms.assetid: 22521ce4-fa0d-b71c-0eaa-d3675dbfc199
ms.date: 06/08/2017
---


# ColorScaleCriterion.Index Property (Excel)

Returns a  **Long** value indicating which threshold the criteria represents. Read-only.


## Syntax

 _expression_ . **Index**

 _expression_ A variable that represents a **ColorScaleCriterion** object.


## Remarks

For a two-color scale conditional format rule, this property will return a value of "1" for the minimum threshold and "2" for the maximum threshold. When using a three-color scale rule, the values will be "1" for the minimum, "2" for the midpoint, and "3" for the maximum thresholds.


## See also


#### Concepts


[ColorScaleCriterion Object](colorscalecriterion-object-excel.md)

