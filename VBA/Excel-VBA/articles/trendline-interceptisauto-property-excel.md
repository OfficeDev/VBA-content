---
title: Trendline.InterceptIsAuto Property (Excel)
keywords: vbaxl10.chm594084
f1_keywords:
- vbaxl10.chm594084
ms.prod: excel
api_name:
- Excel.Trendline.InterceptIsAuto
ms.assetid: ec5ea945-59d7-3ec2-42cd-95c7031880e8
ms.date: 06/08/2017
---


# Trendline.InterceptIsAuto Property (Excel)

 **True** if the point where the trendline crosses the value axis is automatically determined by the regression. Read/write **Boolean** .


## Syntax

 _expression_ . **InterceptIsAuto**

 _expression_ A variable that represents a **Trendline** object.


## Remarks

Setting the  **[Intercept](trendline-interceptisauto-property-excel.md)** property sets this property to **False** .


## Example

This example sets Microsoft Excel to automatically determine the trendline intercept point for Chart1. The example should be run on a 2-D column chart that contains a single series with a trendline.


```vb
Charts("Chart1").SeriesCollection(1).Trendlines(1) _ 
 .InterceptIsAuto = True
```


## See also


#### Concepts


[Trendline Object](trendline-object-excel.md)

