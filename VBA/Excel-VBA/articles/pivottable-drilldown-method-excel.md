---
title: PivotTable.DrillDown Method (Excel)
keywords: vbaxl10.chm235206
f1_keywords:
- vbaxl10.chm235206
ms.prod: excel
ms.assetid: 01824849-6c03-d263-aeb5-68b6c331bf0f
ms.date: 06/08/2017
---


# PivotTable.DrillDown Method (Excel)

Enables you to drill down into the data within an OLAP or PowerPivot based cube hierarchy.


## Syntax

 _expression_ . **DrillDown**_(PivotItem,_ _PivotLine)_

 _expression_ A variable that represents a[PivotTable Object (Excel)](pivottable-object-excel.md) object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
| _PivotItem_|Required|PIVOTITEM|The member from which the drill down is performed.|
| _PivotLine_|Optional|VARIANT|Specifies the line in the PivotTable where the operation starting member resides. In cases where PivotLine is not specified, defaults to the top PivotLine where the member appears.|

### Return value

 **VOID**


## Example

The following sample code demonstrates the  **DrillDown** method as used on a PivotTable.


```vb
ActiveSheet.PivotTables("PivotTable1").DrillDown ActiveSheet.PivotTables( _
      "PivotTable1").PivotFields("[Customer].[Customer Geography].[Country]"). _
      PivotItems("[Customer].[Customer Geography].[Country].&;[Australia]"), _
      ActiveSheet.PivotTables("PivotTable1").PivotRowAxis.PivotLines(1)
```

The following sample code demonstrates the  **DrillDown** method as used on a PivotChart.




```vb
ActiveChart.PivotLayout.PivotTable.DrillDown ActiveChart.PivotLayout.PivotTable _
      .PivotFields("[Customer].[Customer Geography].[Country]").PivotItems( _
      "[Customer].[Customer Geography].[Country].&;[Australia]"), ActiveChart. _
      PivotLayout.PivotTable.PivotRowAxis.PivotLines(1)
```


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)

