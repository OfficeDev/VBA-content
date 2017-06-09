---
title: PivotTable.DisplayEmptyColumn Property (Excel)
keywords: vbaxl10.chm235154
f1_keywords:
- vbaxl10.chm235154
ms.prod: excel
api_name:
- Excel.PivotTable.DisplayEmptyColumn
ms.assetid: 5911c818-282e-bb61-06c2-351cc4c2086d
ms.date: 06/08/2017
---


# PivotTable.DisplayEmptyColumn Property (Excel)

Returns  **True** when the non-empty MDX keyword is included in the query to the OLAP provider for the value axis. The OLAP provider will not return empty columns in the result set. Returns **False** when the non-empty keyword is omitted. Read/write **Boolean** .


## Syntax

 _expression_ . **DisplayEmptyColumn**

 _expression_ A variable that represents a **PivotTable** object.


## Remarks

The PivotTable must be connected to an Online Analytical Processing (OLAP) data source to avoid a run-time error.


## Example

This example determines the display settings for empty columns in a PivotTable. It assumes a PivotTable connected to an OLAP data source exists on the active worksheet.


```vb
Sub CheckSetting() 
 
 Dim pvtTable As PivotTable 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 
 ' Determine display setting for empty columns. 
 If pvtTable.DisplayEmptyColumn = False Then 
 MsgBox "Empty columns will not be displayed." 
 Else 
 MsgBox "Empty columns will be displayed." 
 End If 
 
End Sub
```


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)

