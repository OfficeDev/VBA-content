---
title: Application.ChartDataPointTrack Property (PowerPoint)
keywords: vbapp10.chm502071
f1_keywords:
- vbapp10.chm502071
ms.assetid: c31b3771-d7b1-7559-4480-75f91f1d1f52
ms.date: 06/08/2017
ms.prod: powerpoint
---


# Application.ChartDataPointTrack Property (PowerPoint)

Returns or sets a  **Boolean** that specifies whether charts use cell-reference data-point tracking. Read-write.


## Syntax

 _expression_. **ChartDataPointTrack**

 _expression_ A variable that represents a **Application** object.


### Return Value

 **Boolean**


## Remarks

In cell-reference data-point tracking, data labels track the cell reference that contains the value of the data point, rather than the index number of the data point. Doing so helps to preserve custom formatting applied by the user, even when a chart is sorted or filtered. Setting  **ChartDataPointTrack** to **True** specifies that charts use cell-reference data-point tracking.


## Property value

 **BOOL**


