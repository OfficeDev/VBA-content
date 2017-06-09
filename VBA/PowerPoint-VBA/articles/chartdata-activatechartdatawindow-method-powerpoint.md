---
title: ChartData.ActivateChartDataWindow Method (PowerPoint)
keywords: vbapp10.chm689005
f1_keywords:
- vbapp10.chm689005
ms.assetid: 3364ab9c-ed34-5970-6318-95a694a55354
ms.date: 06/08/2017
ms.prod: powerpoint
---


# ChartData.ActivateChartDataWindow Method (PowerPoint)

Opens a Excel data grid window that contains the full source data for the specified chart.


## Syntax

 _expression_. **ActivateChartDataWindow**

 _expression_ A variable that represents a **ChartData** object.


### Return value

 **VOID**


## Remarks

If the data grid window is already open, this method has no effect.

The  **ActivateChartDataWindow** method differs from the[ChartData.Activate](chartdata-activate-method-powerpoint.md) method in that the former opens the chart in an Excel window within Word, with the Excel ribbon unavailable, whereas the latter opens a full version of Excel, with the ribbon available.


