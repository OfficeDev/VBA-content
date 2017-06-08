---
title: SparklineGroup.PlotBy Property (Excel)
keywords: vbaxl10.chm871091
f1_keywords:
- vbaxl10.chm871091
ms.prod: excel
ms.assetid: 217c6de7-fabf-2642-96a7-aec82f6609a9
ms.date: 06/08/2017
---


# SparklineGroup.PlotBy Property (Excel)

Returns or sets how to plot the sparkline when the data on which it is based is in a square-shaped range. Read/write.


## Syntax

 _expression_ . **PlotBy**

 _expression_ A variable that represents a **[SparklineGroup ](sparklinegroup-object-excel.md)** object.


### Return Value

 **[XlSparklineRowCol](xlsparklinerowcol-enumeration-excel.md)**


## Remarks

This property can only be set if the data on which the sparkline is based is in a square-shaped range, for example if the data is in the range A1:B2. 

The default value for sp data in a square-shaped range is to plot the data by rows ( **xlSparklineRowsSquare** ).


## Property value

 **XLSPARKLINEROWCOL**


## See also


#### Concepts


[SparklineGroup Object](sparklinegroup-object-excel.md)

