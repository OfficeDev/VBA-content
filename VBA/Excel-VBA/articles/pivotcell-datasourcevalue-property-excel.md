---
title: PivotCell.DataSourceValue Property (Excel)
keywords: vbaxl10.chm692087
f1_keywords:
- vbaxl10.chm692087
ms.prod: excel
api_name:
- Excel.PivotCell.DataSourceValue
ms.assetid: 99cd270c-775c-3cca-99dd-1a2864b872b2
ms.date: 06/08/2017
---


# PivotCell.DataSourceValue Property (Excel)

Returns the value last retrieved from the data source for edited cells in a PivotTable report. Read-only


## Syntax

 _expression_ . **DataSourceValue**

 _expression_ A variable that represents a **[PivotCell](pivotcell-object-excel.md)** object.


### Return Value

 **Variant**


## Remarks

Whenever a cell in the values area of a PivotTable report is edited, the  **DataSourceValue** property will hold the value that was last retrieved from the data source before editing took place. For PivotTable report value cells that have not been edited, or for which the data source value has not been explicitly retrieved, this property will return **NULL** . For PivotTable reports with OLAP data sources, the value of the **DataSourceValue** property is retrieved from a separate connection to ensure that it does not contain the value of any writeback operations that the user might have made.

Reading the  **DataSourceValue** property for cells that are outside of the values area of a PivotTable report generates a run-time error.


## See also


#### Concepts


[PivotCell Object](pivotcell-object-excel.md)

