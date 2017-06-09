---
title: PivotTable.RefreshDataSourceValues Method (Excel)
keywords: vbaxl10.chm235194
f1_keywords:
- vbaxl10.chm235194
ms.prod: excel
api_name:
- Excel.PivotTable.RefreshDataSourceValues
ms.assetid: 4312e319-bb90-b8d8-5add-f501553198a6
ms.date: 06/08/2017
---


# PivotTable.RefreshDataSourceValues Method (Excel)

Retrieves the current values from the data source for all edited cells in a PivotTable report that is in writeback mode.


## Syntax

 _expression_ . **RefreshDataSourceValues**

 _expression_ A variable that represents a **[PivotTable](pivottable-object-excel.md)** object.


### Return Value

Nothing


## Remarks

To determine if a PivotTable report is in writeback mode, check the  **[EnableWriteback](pivottable-enablewriteback-property-excel.md)** or **[EnableDataValueEditing](pivottable-enabledatavalueediting-property-excel.md)** properties of the **PivotTable** object, either of which will return **True** . For PivotTable reports that are not in writeback mode, trying to execute this method generates a run-time error.

For PivotTable reports with OLAP data sources, executing the  **RefreshDataSourceValues** method creates a separate connection to the OLAP server and executes the full MDX query (the value of the **PivotTable** . **[MDX](pivottable-mdx-property-excel.md)** property) that is used to perform an update operation to populate the PivotTable report. Excel extracts the values returned for all cells that have been edited in the PivotTable view, and then stores them in the **[DataSourceValue](pivotcell-datasourcevalue-property-excel.md)** property for those cells.

This method applies only to PivotTable reports with OLAP data sources. Trying to execute this method or PivotTable reports with non-OLAP data sources generates a run-time error.


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)

