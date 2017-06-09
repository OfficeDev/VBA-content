---
title: Application.ChartDataPointTrack Property (Excel)
keywords: vbaxl10.chm133341
f1_keywords:
- vbaxl10.chm133341
ms.prod: excel
ms.assetid: 124b4d82-de33-c5df-7aa0-1a9c3484a680
ms.date: 06/08/2017
---


# Application.ChartDataPointTrack Property (Excel)

 **True** will cause all charts in newly created documents to use the cell reference tracking behavior. **Boolean**


## Syntax

 _expression_ . **ChartDataPointTrack**

 _expression_ A variable that represents a **Application** object.


## Notes

Data labels will now track the  _actual_ data point to which they are attached (as opposed to the legacy behavior of tracking the _index_ of the data point), allowing the label-to-point relationship to persist across events such as filter and sort.


## Property value

 **BOOL**


## See also


#### Concepts


[Application Object](application-object-excel.md)

