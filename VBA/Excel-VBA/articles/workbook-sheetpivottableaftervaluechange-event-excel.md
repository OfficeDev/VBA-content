---
title: Workbook.SheetPivotTableAfterValueChange Event (Excel)
keywords: vbaxl10.chm503102
f1_keywords:
- vbaxl10.chm503102
ms.prod: excel
api_name:
- Excel.Workbook.SheetPivotTableAfterValueChange
ms.assetid: 8460f5f1-d415-7aac-6a3d-fa0944036e9c
ms.date: 06/08/2017
---


# Workbook.SheetPivotTableAfterValueChange Event (Excel)

Occurs after a cell or range of cells inside a PivotTable are edited or recalculated (for cells that contain formulas).


## Syntax

 _expression_ . **SheetPivotTableAfterValueChange**( **_Sh_** , **_TargetPivotTable_** , **_TargetRange_** )

 _expression_ A variable that represents a **[Workbook](workbook-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Sh_|Required| **Object**|The worksheet that contains the PivotTable.|
| _TargetPivotTable_|Required| **[PivotTable](pivottable-object-excel.md)**|The PivotTable that contains the edited or recalculated cells.|
| _TargetRange_|Required| **[Range](range-object-excel.md)**|The range that contains all the edited or recalcuated cells.|

### Return Value

 **Nothing**


## Remarks

The  **PivotTableAfterValueChange** event does not occur under any conditions other than editing or recalculating cells. For example, it will not occur when the PivotTable is refreshed, sorted, filtered, or drilled down on, even though those operations move cells and potentially retrieve new values from the OLAP data source.


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

