---
title: Trendline.NameIsAuto Property (Excel)
keywords: vbaxl10.chm594086
f1_keywords:
- vbaxl10.chm594086
ms.prod: excel
api_name:
- Excel.Trendline.NameIsAuto
ms.assetid: 4e14cc52-a9f5-3dda-8be9-7afd97d79583
ms.date: 06/08/2017
---


# Trendline.NameIsAuto Property (Excel)

 **True** if Microsoft Excel automatically determines the name of the trendline. Read/write **Boolean** .


## Syntax

 _expression_ . **NameIsAuto**

 _expression_ A variable that represents a **Trendline** object.


## Example

This example sets Microsoft Excel to automatically determine the name for trendline one in Chart1. The example should be run on a 2-D column chart that contains a single series with a trendline.


```vb
Charts("Chart1").SeriesCollection(1) _ 
 .Trendlines(1).NameIsAuto = True
```


## See also


#### Concepts


[Trendline Object](trendline-object-excel.md)

