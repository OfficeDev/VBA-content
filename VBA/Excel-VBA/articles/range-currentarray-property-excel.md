---
title: Range.CurrentArray Property (Excel)
keywords: vbaxl10.chm144110
f1_keywords:
- vbaxl10.chm144110
ms.prod: excel
api_name:
- Excel.Range.CurrentArray
ms.assetid: 147f8834-5aef-900f-75de-df91a6a76005
ms.date: 06/08/2017
---


# Range.CurrentArray Property (Excel)

If the specified cell is part of an array, returns a  **[Range](range-object-excel.md)** object that represents the entire array. Read-only.


## Syntax

 _expression_ . **CurrentArray**

 _expression_ A variable that represents a **Range** object.


## Example

This example assumes that cell A1 on Sheet1 is the active cell and that the active cell is part of an array that includes cells A1:A10. The example selects cells A1:A10 on Sheet1.


```vb
ActiveCell.CurrentArray.Select
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

