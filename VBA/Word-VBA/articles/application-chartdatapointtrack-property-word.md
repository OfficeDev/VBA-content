---
title: Application.ChartDataPointTrack Property (Word)
keywords: vbawd10.chm158335470
f1_keywords:
- vbawd10.chm158335470
ms.prod: word
ms.assetid: dea8365d-aadf-6667-ade0-2bef1622fd39
ms.date: 06/08/2017
---


# Application.ChartDataPointTrack Property (Word)

Returns or sets a  **Boolean** that specifies whether charts use cell-reference data-point tracking. Read-write.


## Syntax

 _expression_ . **ChartDataPointTrack**

 _expression_ A variable that represents a **Application** object.


## Remarks

In cell-reference data-point tracking, data labels track the cell reference that contains the value of the data point, rather than the index number of the data point. Doing so helps to preserve custom formatting applied by the user, even when a chart is sorted or filtered. Setting  **ChartDataPointTrack** to **True** specifies that charts use cell-reference data-point tracking.


## Property value

 **BOOL**


## See also


#### Concepts


[Application Object](application-object-word.md)

