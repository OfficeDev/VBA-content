---
title: XlCellInsertionMode Enumeration (Excel)
ms.prod: excel
api_name:
- Excel.XlCellInsertionMode
ms.assetid: 582f504f-8acf-c359-186e-35429192b6b0
ms.date: 06/08/2017
---


# XlCellInsertionMode Enumeration (Excel)

Specifies the way that rows on the specified worksheet are added or deleted to accommodate the number of rows in a recordset returned by a query.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **xlInsertDeleteCells**|1|Partial rows are inserted or deleted to match the exact number of rows required for the new recordset.|
| **xlInsertEntireRows**|2|Entire rows are inserted, if necessary, to accommodate any overflow. No cells or rows are deleted from the worksheet.|
| **xlOverwriteCells**|0|No new cells or rows are added to the worksheet. Data in surrounding cells is overwritten to accommodate any overflow.|

