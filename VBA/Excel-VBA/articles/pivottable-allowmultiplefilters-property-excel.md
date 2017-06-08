---
title: PivotTable.AllowMultipleFilters Property (Excel)
keywords: vbaxl10.chm235178
f1_keywords:
- vbaxl10.chm235178
ms.prod: excel
api_name:
- Excel.PivotTable.AllowMultipleFilters
ms.assetid: e6e39932-9d20-d34b-a2b1-6b34e4bfb270
ms.date: 06/08/2017
---


# PivotTable.AllowMultipleFilters Property (Excel)

Sets or retrieves a value that indicates whether a PivotField can have multiple filters applied to it at the same time. Read/write  **Boolean** .


## Syntax

 _expression_ . **AllowMultipleFilters**

 _expression_ A variable that represents a **PivotTable** object.


## Remarks

Default value is  **False** .

When this property is set to  **True** , multiple filters can be applied to single PivotFields. When it is set to **False** , applying a filter to a PivotField that is already filtered will remove the existing filter and apply the new one. Setting this property to **False** when the PivotTable has fields with more than one filter applied will silently remove all filters in the PivotTable without displaying any alert. However, there is an alert when this is done through the user interface.


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)

