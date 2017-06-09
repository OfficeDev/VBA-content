---
title: PivotTable.PrintTitles Property (Excel)
keywords: vbaxl10.chm235131
f1_keywords:
- vbaxl10.chm235131
ms.prod: excel
api_name:
- Excel.PivotTable.PrintTitles
ms.assetid: a8138146-bfe9-1af9-c101-0c095c4a91a5
ms.date: 06/08/2017
---


# PivotTable.PrintTitles Property (Excel)

 **True** if the print titles for the worksheet are set based on the PivotTable report. **False** if the print titles for the worksheet are used. The default value is **False** . Read/write **Boolean** .


## Syntax

 _expression_ . **PrintTitles**

 _expression_ A variable that represents a **PivotTable** object.


## Remarks

The row print titles are set to the rows that contain the PivotTable report's column field items. The column print titles are set to the columns that contain the row items.


## Example

This example specifies that the print title set for the worksheet is printed when the fourth PivotTable report on the active worksheet is printed.


```vb
ActiveSheet.PivotTables("PivotTable4").PrintTitles = True
```


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)

