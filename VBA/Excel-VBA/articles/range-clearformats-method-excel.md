---
title: Range.ClearFormats Method (Excel)
keywords: vbaxl10.chm144096
f1_keywords:
- vbaxl10.chm144096
ms.prod: excel
api_name:
- Excel.Range.ClearFormats
ms.assetid: 37777379-857a-f4c7-86aa-b109d5f25757
ms.date: 06/08/2017
---


# Range.ClearFormats Method (Excel)

Clears the formatting of the object.


## Syntax

 _expression_ . **ClearFormats**

 _expression_ A variable that represents a **Range** object.


### Return Value

Variant


## Example

This example clears all formatting from cells A1:G37 on Sheet1.


```vb
Worksheets("Sheet1").Range("A1:G37").ClearFormats
```

This example clears the formatting from embedded chart one on Sheet1.




```vb
Worksheets("Sheet1").ChartObjects(1).Chart.ChartArea.ClearFormats
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

