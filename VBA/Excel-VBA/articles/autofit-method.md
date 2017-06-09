---
title: AutoFit Method
keywords: vbagr10.chm65773
f1_keywords:
- vbagr10.chm65773
ms.prod: excel
api_name:
- Excel.AutoFit
ms.assetid: 45dea7dd-7695-1f72-9bf7-9ab4cbbd74ec
ms.date: 06/08/2017
---


# AutoFit Method

Changes the width of the columns in the specified range to achieve the best fit.

 _expression_. **AutoFit**

 _expression_ Required. An expression that returns a **Range** object. Must be a row or a range of rows, or a column or a range of columns. Otherwise, this method causes an error.


## Remarks

One unit of column width is equal to the width of one character in the Normal style.


## Example

This example changes the width of columns A through I on the datasheet to achieve the best fit.


```
myChart.Application.DataSheet.Columns("A:I").AutoFit
```

This example changes the width of columns A through E on the datasheet to achieve the best fit, based only on the contents of cells A1:E1.




```
myChart.Application.DataSheet.Range("A1:E1").Columns.AutoFit
```


