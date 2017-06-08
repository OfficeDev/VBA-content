---
title: PivotCache.UpgradeOnRefresh Property (Excel)
keywords: vbaxl10.chm227109
f1_keywords:
- vbaxl10.chm227109
ms.prod: excel
api_name:
- Excel.PivotCache.UpgradeOnRefresh
ms.assetid: 9110a82b-9ac7-3d9e-8386-827cd828aace
ms.date: 06/08/2017
---


# PivotCache.UpgradeOnRefresh Property (Excel)

Contains information on whether to upgrade the PivotCache and all connected PivotTables on the next refresh. Read/write  **Boolean** .


## Syntax

 _expression_ . **UpgradeOnRefresh**

 _expression_ A variable that represents a **PivotCache** object.


## Remarks

The default value is  **False** . If the property is set to **True** for a PivotCache, refreshing any PivotTable attached to that PivotCache will upgrade the PivotCache and all the attached PivotTables to **xlPivotTableVersion12** (PivotTable.Version = 3) as part of the refresh.

If the property is set to  **False** for a PivotCache, refreshing any PivotTable attached to that PivotCache will not change the version of the PivotCache, nor the version of all the attached PivotTables. They all stay the same version as before the refresh.

Saving to an Excel 2007 or later file format, when in compatibility mode, will set this property to  **True** for all PivotCaches in the workbook.


## See also


#### Concepts


[PivotCache Object](pivotcache-object-excel.md)

