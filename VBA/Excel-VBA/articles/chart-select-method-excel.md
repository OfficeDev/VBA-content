---
title: Chart.Select Method (Excel)
keywords: vbaxl10.chm148094
f1_keywords:
- vbaxl10.chm148094
ms.prod: excel
api_name:
- Excel.Chart.Select
ms.assetid: 20f866f4-14b9-075c-372c-47a9f536f0c3
ms.date: 06/08/2017
---


# Chart.Select Method (Excel)

Selects the object.


## Syntax

 _expression_ . **Select**( **_Replace_** )

 _expression_ A variable that represents a **Chart** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Replace_|Optional| **Variant**| (used only with sheets). **True** to replace the current selection with the specified object. **False** to extend the current selection to include any previously selected objects and the specified object.|

## Remarks

To select a cell or a range of cells, use the  **Select** method. To make a single cell the active cell, use the **[Activate](chart-activate-method-excel.md)** method.


## See also


#### Concepts


[Chart Object](chart-object-excel.md)

