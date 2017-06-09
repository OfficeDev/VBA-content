---
title: ChartObjects.ProtectChartObject Property (Excel)
keywords: vbaxl10.chm497098
f1_keywords:
- vbaxl10.chm497098
ms.prod: excel
api_name:
- Excel.ChartObjects.ProtectChartObject
ms.assetid: e0685fbd-84a5-36c4-a5ab-06127937f2c8
ms.date: 06/08/2017
---


# ChartObjects.ProtectChartObject Property (Excel)

 **True** if the embedded chart frame cannot be moved, resized, or deleted through the user interface. Read/write **Boolean** .


## Syntax

 _expression_ . **ProtectChartObject**

 _expression_ A variable that represents a **ChartObjects** object.


## Remarks

Setting this property to  **True** will not protect the embedded chart frame from being modified through the object model.


## Example


```vb
Worksheets(1).ChartObjects(1).ProtectChartObject = True
```


## See also


#### Concepts


[ChartObjects Object](chartobjects-object-excel.md)

