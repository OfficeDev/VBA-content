---
title: Document.ChartDataPointTrack Property (Word)
keywords: vbawd10.chm158007914
f1_keywords:
- vbawd10.chm158007914
ms.prod: word
ms.assetid: 3b9bb881-4e9b-d8bc-dc57-4a4be573a5a0
ms.date: 06/08/2017
---


# Document.ChartDataPointTrack Property (Word)

Returns or sets a  **Boolean** that specifies whether charts in the active document use cell-reference data-point tracking. Read-write.


## Syntax

 _expression_ . **ChartDataPointTrack**

 _expression_ A variable that represents a **Document** object.


## Remarks

In cell-reference data-point tracking, data labels track the cell reference that contains the value of the data point, rather than the index number of the data point. Doing so helps to preserve custom formatting applied by the user, even when a chart is sorted or filtered. Setting  **ChartDataPointTrack** to **True** specifies that charts in the active document use cell-reference data-point tracking.


## Property value

 **BOOL**


## See also


#### Concepts


[Document Object](document-object-word.md)

