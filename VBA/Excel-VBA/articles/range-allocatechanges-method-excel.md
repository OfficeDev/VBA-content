---
title: Range.AllocateChanges Method (Excel)
keywords: vbaxl10.chm144253
f1_keywords:
- vbaxl10.chm144253
ms.prod: excel
api_name:
- Excel.Range.AllocateChanges
ms.assetid: c751c5fb-ce22-64d1-669c-fdb064cf0408
ms.date: 06/08/2017
---


# Range.AllocateChanges Method (Excel)

Performs a writeback operation for all edited cells in a range based on an OLAP data source.


## Syntax

 _expression_ . **AllocateChanges**

 _expression_ A variable that represents a **[Range](range-object-excel.md)** object.


## Remarks

The  **AllocateChanges** method will execute an **UPDATE CUBE** statement for all changes made in the range since the last apply changes operation was committed. This method generates a run-time error if it is executed on a range based on a non-OLAP data source.


## See also


#### Concepts


[Range Object](range-object-excel.md)

