---
title: Application.SheetPivotTableAfterValueChange Event (Excel)
keywords: vbaxl10.chm504104
f1_keywords:
- vbaxl10.chm504104
ms.prod: excel
api_name:
- Excel.Application.SheetPivotTableAfterValueChange
ms.assetid: 07cab356-1a13-a839-7344-a4de99dba55e
ms.date: 06/08/2017
---


# Application.SheetPivotTableAfterValueChange Event (Excel)

Occurs after a cell or range of cells inside a PivotTable are edited or recalculated (for cells that contain formulas).


## Syntax

 _expression_ . **SheetPivotTableAfterValueChange**( **_Sh_** , **_TargetPivotTable_** , **_TargetRange_** )

 _expression_ A variable that represents a **[Application](application-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Sh_|Required| **Object**|The worksheet that contains the PivotTable|
| _TargetPivotTable_|Required| **[PivotTable](pivottable-object-excel.md)**|The PivotTable that contains the edited or recalculated cells.|
| _TargetRange_|Required| **[Range](range-object-excel.md)**|The range that contains all the edited or recalcuated cells.|

### Return Value

 **Nothing**


## Remarks

The  **PivotTableAfterValueChange** event does not occur under any conditions other than editing or recalculating cells. For example, it will not occur when the PivotTable is refreshed, sorted, filtered, or drilled down on, even though those operations move cells and potentially retrieve new values from the OLAP data source.


## See also


#### Concepts


[Application Object](application-object-excel.md)

