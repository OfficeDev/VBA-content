---
title: ChartObject.SendToBack Method (Excel)
keywords: vbaxl10.chm494091
f1_keywords:
- vbaxl10.chm494091
ms.prod: excel
api_name:
- Excel.ChartObject.SendToBack
ms.assetid: a8f0f721-15ba-662f-ac17-0ac1657e3413
ms.date: 06/08/2017
---


# ChartObject.SendToBack Method (Excel)

Sends the object to the back of the z-order.


## Syntax

 _expression_ . **SendToBack**

 _expression_ A variable that represents a **ChartObject** object.


### Return Value

Variant


## Example

This example sends embedded chart one on Sheet1 to the back of the z-order.


```vb
Worksheets("Sheet1").ChartObjects(1).SendToBack
```


## See also


#### Concepts


[ChartObject Object](chartobject-object-excel.md)

