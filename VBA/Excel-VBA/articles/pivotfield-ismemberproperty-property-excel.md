---
title: PivotField.IsMemberProperty Property (Excel)
keywords: vbaxl10.chm240131
f1_keywords:
- vbaxl10.chm240131
ms.prod: excel
api_name:
- Excel.PivotField.IsMemberProperty
ms.assetid: e24e6e84-2c27-5d33-78c4-b48e96d48e5d
ms.date: 06/08/2017
---


# PivotField.IsMemberProperty Property (Excel)

Returns  **True** when the PivotField contains member properties. Read-only **Boolean** .


## Syntax

 _expression_ . **IsMemberProperty**

 _expression_ A variable that represents a **PivotField** object.


## Remarks

This property will return a run-time error if an Online Analytical Processing (OLAP) data source is not used.


## Example

This example determines if the PivotTable field contains member properties and notifies the user. It assumes that a PivotTable exists on the active worksheet and that it is connected to an OLAP data source.


```vb
Sub CheckForMembers() 
 
 Dim pvtTable As PivotTable 
 Dim pvtField As PivotField 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 Set pvtField = pvtTable.PivotFields(1) 
 
 ' Determine if member properties exist and notify user. 
 If pvtField.IsMemberProperty = True Then 
 MsgBox "The PivotField contains member properties." 
 Else 
 MsgBox "The PivotField does not contain member properties." 
 End If 
 
End Sub
```


## See also


#### Concepts


[PivotField Object](pivotfield-object-excel.md)

