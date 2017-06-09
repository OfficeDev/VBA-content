---
title: Worksheet.PivotTableAfterValueChange Event (Excel)
keywords: vbaxl10.chm502082
f1_keywords:
- vbaxl10.chm502082
ms.prod: excel
api_name:
- Excel.Worksheet.PivotTableAfterValueChange
ms.assetid: 097e1c1e-4df6-a0d1-de67-0e0752d2286a
ms.date: 06/08/2017
---


# Worksheet.PivotTableAfterValueChange Event (Excel)

Occurs after a cell or range of cells inside a PivotTable are edited or recalculated (for cells that contain formulas).


## Syntax

 _expression_ . **PivotTableAfterValueChange**( **_TargetPivotTable_** , **_TargetRange_** )

 _expression_ A variable that represents a **[Worksheet](worksheet-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _TargetPivotTable_|Required| **[PivotTable](pivottable-object-excel.md)**|The PivotTable that contains the edited or recalculated cells.|
| _TargetRange_|Required| **[Range](range-object-excel.md)**|The range that contains all the edited or recalcuated cells.|

### Return Value

Nothing


## Remarks

The  **PivotTableAfterValueChange** event does not occur under any conditions other than editing or recalculating cells. For example, it will not occur when the PivotTable is refreshed, sorted, filtered, or drilled down on, even though those operations move cells and potentially retrieve new values from the OLAP data source.


## See also


#### Concepts


[Worksheet Object](worksheet-object-excel.md)

