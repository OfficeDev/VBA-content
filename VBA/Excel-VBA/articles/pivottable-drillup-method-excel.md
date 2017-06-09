---
title: PivotTable.DrillUp Method (Excel)
keywords: vbaxl10.chm235207
f1_keywords:
- vbaxl10.chm235207
ms.prod: excel
ms.assetid: 18933878-53c5-ef64-afe7-919b0a1564f8
ms.date: 06/08/2017
---


# PivotTable.DrillUp Method (Excel)

Enables you to drill up into the data within an OLAP or PowerPivot based cube hierarchy.


## Syntax

 _expression_ . **DrillUp**_(PivotItem,_ _PivotLine,_ _LevelUniqueName)_

 _expression_ A variable that represents a **PivotTable** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
| _PivotItem_|Required|PIVOTITEM|The member from which the drill up is performed.|
| _PivotLine_|Optional|VARIANT|Specifies the line in the PivotTable where the operation starting member resides. In cases where PivotLine is not specified, defaults to the top PivotLine where the member appears.|
| _LevelUniqueName_|Optional|VARIANT|The target for a multi-level drill up. The default action if not specified is a one level drill up.|

### Return value

 **VOID**


### Examples

The following sample code demonstrates a single-level drill up on a PivotTable


```vb
ActiveSheet.PivotTables("PivotTable1").DrillUp ActiveSheet.PivotTables( _
      "PivotTable1").PivotFields("[Customer].[Customer Geography].[Postal Code]"). _
      PivotItems( _
      "[Customer].[Customer Geography].[Postal Code].&;[2450]&;[Coffs Harbour]"), _
      ActiveSheet.PivotTables("PivotTable1").PivotRowAxis.PivotLines(1)
```

The following sample code demonstrates a level drill up on a PivotChart




```vb
ActiveChart.PivotLayout.PivotTable.DrillUp ActiveChart.PivotLayout.PivotTable. _
      PivotFields("[Customer].[Customer Geography].[Postal Code]").PivotItems( _
      "[Customer].[Customer Geography].[Postal Code].&;[2450]&;[Coffs Harbour]"), _
      ActiveChart.PivotLayout.PivotTable.PivotRowAxis.PivotLines(1)
```

The following sample code demonstrates multi-level drill up on a PivotTable




```vb
ActiveSheet.PivotTables("PivotTable1").DrillUp ActiveSheet.PivotTables( _
     "PivotTable1").PivotFields("[Customer].[Customer Geography].[City]").PivotItems _
     ("[Customer].[Customer Geography].[City].&;[Coffs Harbour]&;[NSW]"), ActiveSheet. _
     PivotTables("PivotTable1").PivotRowAxis.PivotLines(1), _
     "[Customer].[Customer Geography].[Country]"
```

The following sample code demonstrates multi-level drill up on a PivotChart




```vb
ActiveChart.PivotLayout.PivotTable.DrillUp ActiveChart.PivotLayout.PivotTable. _
     PivotFields("[Customer].[Customer Geography].[Postal Code]").PivotItems( _
     "[Customer].[Customer Geography].[Postal Code].&;[2450]&;[Coffs Harbour]"), _
     ActiveChart.PivotLayout.PivotTable.PivotRowAxis.PivotLines(1) , _
     "[Customer].[Customer Geography].[Country]"
```


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)

