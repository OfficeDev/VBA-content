---
title: PivotTable.Allocation Property (Excel)
keywords: vbaxl10.chm235187
f1_keywords:
- vbaxl10.chm235187
ms.prod: excel
api_name:
- Excel.PivotTable.Allocation
ms.assetid: ac7bd537-97f0-f643-3e34-dd13e49ac149
ms.date: 06/08/2017
---


# PivotTable.Allocation Property (Excel)

Returns or sets whether to run an  **UPDATE CUBE** statement for each cell is edited, or only when the user chooses to calculate changes when performing what-if analysis on a PivotTable based on an OLAP data source. Read/write


## Syntax

 _expression_ . **Allocation**

 _expression_ A variable that represents a **[PivotTable](pivottable-object-excel.md)** object.


### Return Value

 **[XlAllocation](xlallocation-enumeration-excel.md)**


## Remarks

The  **Allocation** property corresponds to the **Calculate with changes** setting in the **What-If Analysis Settings** dialog box. The default setting is **xlManualAllocation** , which corresponds to the **Manually (when selecting calculate PivotTable with changes)** setting.


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)

