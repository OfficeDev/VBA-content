---
title: Range.UseStandardHeight Property (Excel)
keywords: vbaxl10.chm144213
f1_keywords:
- vbaxl10.chm144213
ms.prod: excel
api_name:
- Excel.Range.UseStandardHeight
ms.assetid: 59e0be39-25ea-c18d-919d-506d4f041f45
ms.date: 06/08/2017
---


# Range.UseStandardHeight Property (Excel)

 **True** if the row height of the **Range** object equals the standard height of the sheet. Returns **Null** if the range contains more than one row and the rows aren't all the same height. Read/write **Variant** .


## Syntax

 _expression_ . **UseStandardHeight**

 _expression_ A variable that represents a **Range** object.


## Example

This example sets the height of row one on Sheet1 to the standard height.


```vb
Worksheets("Sheet1").Rows(1).UseStandardHeight = True
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

