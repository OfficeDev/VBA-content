---
title: Range.Style Property (Excel)
keywords: vbaxl10.chm144204
f1_keywords:
- vbaxl10.chm144204
ms.prod: excel
api_name:
- Excel.Range.Style
ms.assetid: 78c536c9-7fda-3171-2a93-5c4e57bb8207
ms.date: 06/08/2017
---


# Range.Style Property (Excel)

Returns or sets a  **Variant** value, containing a **[Style](style-object-excel.md)** object, that represents the style of the specified range.


## Syntax

 _expression_ . **Style**

 _expression_ A variable that represents a **Range** object.


## Example

This example applies the Normal style to cell A1 on Sheet1.


```vb
Worksheets("Sheet1").Range("A1").Style.Name = "Normal"
```

If cell B4 on Sheet1 currently has the Normal style applied, this example applies the Percent style.




```vb
If Worksheets("Sheet1").Range("B4").Style.Name = "Normal" Then 
 Worksheets("Sheet1").Range("B4").Style.Name = "Percent" 
End If
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

