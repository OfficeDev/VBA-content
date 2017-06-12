---
title: PivotCell.DiscardChange Method (Excel)
keywords: vbaxl10.chm692086
f1_keywords:
- vbaxl10.chm692086
ms.prod: excel
api_name:
- Excel.PivotCell.DiscardChange
ms.assetid: 26bd8db6-05c9-791a-be69-2f141053c746
ms.date: 06/08/2017
---


# PivotCell.DiscardChange Method (Excel)

Discards changes to the specified cell in a PivotTable report.


## Syntax

 _expression_ . **DiscardChange**

 _expression_ A variable that represents a **[PivotCell](pivotcell-object-excel.md)** object.


### Return Value

 **Nothing**


## Remarks

For a cell in a PivotTable based on an OLAP data source, the  **DiscardChange** method removes the value or formula entered in the specified cell, and then runs a PivotTable update to retrieve the latest value from the data source. The data source value will be set to **NULL** by this method for the specified cell. The method also removes all changes made to this cell from the Excel change list so that it will no longer be included in the **UPDATE CUBE** statement for the data source. This method also recreates the session state for all other changes made in the transaction. This includes executing a **ROLLBACK TRANSACTION** statement, and then executing an **UPDATE CUBE** statement with all changes except for the specified cell. For a cell in a PivotTable based on a non-OLAP data source, this method clears the edited cell.


## See also


#### Concepts


[PivotCell Object](pivotcell-object-excel.md)

