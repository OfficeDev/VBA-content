---
title: PivotTable.CalculatedMembers Property (Excel)
keywords: vbaxl10.chm235145
f1_keywords:
- vbaxl10.chm235145
ms.prod: excel
api_name:
- Excel.PivotTable.CalculatedMembers
ms.assetid: 65e7ffd6-e01d-f8fc-3adb-a1bcb1046fcf
ms.date: 06/08/2017
---


# PivotTable.CalculatedMembers Property (Excel)

Returns a  **[CalculatedMembers](calculatedmembers-object-excel.md)** collection representing all the calculated members and calculated measures for an OLAP PivotTable.


## Syntax

 _expression_ . **CalculatedMembers**

 _expression_ A variable that represents a **PivotTable** object.


## Remarks

This property is used for Online Analytical Processing (OLAP) sources; a non-OLAP source will return a run-time error.


## Example

This example adds a set to the PivotTable. It assumes a PivotTable exists on the active worksheet that is connected to an OLAP data source which contains a field titled "[Product].[All Products]".


```vb
Sub UseCalculatedMember() 
 
 Dim pvtTable As PivotTable 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 
 ' Add the calculated member. 
 pvtTable.CalculatedMembers.Add Name:="[Beef]", _ 
 Formula:="'{[Product].[All Products].Children}'", _ 
 Type:=xlCalculatedSet 
 
End Sub
```


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)

