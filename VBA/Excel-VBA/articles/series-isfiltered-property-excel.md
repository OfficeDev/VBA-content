---
title: Series.IsFiltered Property (Excel)
keywords: vbaxl10.chm578128
f1_keywords:
- vbaxl10.chm578128
ms.prod: excel
ms.assetid: 90c1564c-439c-de1f-8082-0de3372c0566
ms.date: 06/08/2017
---


# Series.IsFiltered Property (Excel)

This setting controls whether the series has been filtered out from the chart. The default value is  **False** . **Boolean** Read/Write.


## Syntax

 _expression_ . **IsFiltered**

 _expression_ A variable that represents a **Series** object.


## Remarks

When a user filters out a series, the series  **IsFiltered** property switches to **True** , and the series is transferred out of its parent **SeriesCollection** . However, that series still remains in its parent **FullSeriesCollection** . When a user filters the series back in, it is inserted back in its original place in the **SeriesCollection** .


## Property value

 **BOOL**


## See also


#### Concepts


[Series Object](series-object-excel.md)

