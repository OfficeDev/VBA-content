---
title: ChartObject.ProtectChartObject Property (Excel)
keywords: vbaxl10.chm494100
f1_keywords:
- vbaxl10.chm494100
ms.prod: excel
api_name:
- Excel.ChartObject.ProtectChartObject
ms.assetid: 0fd7830a-5c07-89f4-190d-b4b231512de7
ms.date: 06/08/2017
---


# ChartObject.ProtectChartObject Property (Excel)

 **True** if the embedded chart frame cannot be moved, resized, or deleted through the user interface. Read/write **Boolean** .


## Syntax

 _expression_ . **ProtectChartObject**

 _expression_ A variable that represents a **ChartObject** object.


## Remarks

Setting this property to  **True** will not protect the embedded chart frame from being modified through the object model.


## Example

This example protects embedded chart one on worksheet one.


```vb
Worksheets(1).ChartObjects(1).ProtectChartObject = True
```


## See also


#### Concepts


[ChartObject Object](chartobject-object-excel.md)

