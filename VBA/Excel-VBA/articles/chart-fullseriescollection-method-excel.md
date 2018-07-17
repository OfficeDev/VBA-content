---
title: Chart.FullSeriesCollection Method (Excel)
keywords: vbaxl10.chm149194
f1_keywords:
- vbaxl10.chm149194
ms.prod: excel
ms.assetid: 875c18cf-064f-6b2f-2650-f5d07c16bc4d
ms.date: 06/08/2017
---


# Chart.FullSeriesCollection Method (Excel)

Enables retrieving the filtered out series specified by the Index argument.


## Syntax

 _expression_ . **FullSeriesCollection**_(Index)_

 _expression_ A variable that represents a **Chart** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Optional|VARIANT|The indexed number of the filtered out  **Series** object.|

### Return value

 **OBJECT**


## Remarks

 **Series** objects in hidden rows or columns do not appear in the current series collection unless the user has enabled the **Show data in hidden rows and columns** option in the **Select Data** dialog.


## See also


#### Concepts


[Chart Object](chart-object-excel.md)

