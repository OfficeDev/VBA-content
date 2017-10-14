---
title: PivotTable.DisplayEmptyRow Property (Excel)
keywords: vbaxl10.chm235153
f1_keywords:
- vbaxl10.chm235153
ms.prod: excel
api_name:
- Excel.PivotTable.DisplayEmptyRow
ms.assetid: c1e20ff1-44db-47a0-8e4b-7db7d2ad7cb2
ms.date: 06/08/2017
---


# PivotTable.DisplayEmptyRow Property (Excel)

Returns  **True** when the non-empty MDX keyword is included in the query to the OLAP provider for the category axis. The OLAP provider will not return empty rows in the result set. Returns **False** when the non-empty keyword is omitted. Read/write **Boolean** .


## Syntax

 _expression_ . **DisplayEmptyRow**

 _expression_ A variable that represents a **PivotTable** object.


## Remarks

The PivotTable must be connected to an Online Analytical Processing (OLAP) data source to avoid a run-time error.


## Example

This example determines the display settings for empty rows in a PivotTable. It assumes a PivotTable connected to an OLAP data source exists on the active worksheet.


```vb
Sub CheckSetting() 
 
 Dim pvtTable As PivotTable 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 
 ' Determine display setting for empty rows. 
 If pvtTable.DisplayEmptyRow = False Then 
 MsgBox "Empty rows will not be displayed." 
 Else 
 MsgBox "Empty rows will be displayed." 
 End If 
 
End Sub
```


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)

