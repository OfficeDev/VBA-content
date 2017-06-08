---
title: PivotTable.DataPivotField Property (Excel)
keywords: vbaxl10.chm235140
f1_keywords:
- vbaxl10.chm235140
ms.prod: excel
api_name:
- Excel.PivotTable.DataPivotField
ms.assetid: 00b62ffd-76bd-cd4b-218c-b6d695150efb
ms.date: 06/08/2017
---


# PivotTable.DataPivotField Property (Excel)

Returns a  **[PivotField](pivotfield-object-excel.md)** object that represents all the data fields in a PivotTable. Read-only.


## Syntax

 _expression_ . **DataPivotField**

 _expression_ A variable that represents a **PivotTable** object.


## Example

This example moves the second  **PivotItem** object to the first position. It assumes a PivotTable exists on the active worksheet and that the PivotTable contains data fields.


```vb
Sub UseDataPivotField() 
 
 Dim pvtTable As PivotTable 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 
 ' Move second PivotItem to the first position in PivotTable. 
 pvtTable.DataPivotField.PivotItems(2).Position = 1 
 
End Sub
```


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)

