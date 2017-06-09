---
title: DataSheet Property
keywords: vbagr10.chm5207292
f1_keywords:
- vbagr10.chm5207292
ms.prod: excel
api_name:
- Excel.DataSheet
ms.assetid: d7ccd394-e9b7-2967-76a4-60e5dda40a84
ms.date: 06/08/2017
---


# DataSheet Property

Returns the  **[DataSheet](datasheet-object.md)** object. Read-only.


## Example

This example sets the value of cell A1 on the datasheet to 3.14159.


```vb
With myChart.Application 
 .DataSheet.Range("A1").Value = 3.14159 
End With
```


