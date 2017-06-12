---
title: ChartObject.RoundedCorners Property (Excel)
keywords: vbaxl10.chm494101
f1_keywords:
- vbaxl10.chm494101
ms.prod: excel
api_name:
- Excel.ChartObject.RoundedCorners
ms.assetid: cb58389a-0235-384e-e32a-e669e789bacc
ms.date: 06/08/2017
---


# ChartObject.RoundedCorners Property (Excel)

 **True** if the embedded chart has rounded corners. Read/write **Boolean** .


## Syntax

 _expression_ . **RoundedCorners**

 _expression_ A variable that represents a **ChartObject** object.


## Example

This example adds rounded corners to embedded chart one on Sheet1.


```vb
Worksheets("Sheet1").ChartObjects(1).RoundedCorners = True
```


## See also


#### Concepts


[ChartObject Object](chartobject-object-excel.md)

