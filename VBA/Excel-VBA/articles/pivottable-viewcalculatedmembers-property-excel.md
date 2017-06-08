---
title: PivotTable.ViewCalculatedMembers Property (Excel)
keywords: vbaxl10.chm235144
f1_keywords:
- vbaxl10.chm235144
ms.prod: excel
api_name:
- Excel.PivotTable.ViewCalculatedMembers
ms.assetid: 2d1f752a-0bab-baa6-a9b0-e158cc9a4f09
ms.date: 06/08/2017
---


# PivotTable.ViewCalculatedMembers Property (Excel)

When set to  **True** (default), calculated members for Online Analytical Processing (OLAP) PivotTables can be viewed. Read/write **Boolean** .


## Syntax

 _expression_ . **ViewCalculatedMembers**

 _expression_ A variable that represents a **PivotTable** object.


## Example

This example determines if calculated members can be viewed on the PivotTable and notifies the user. It assumes that a PivotTable exists on the active worksheet.


```vb
Sub CheckViewCalculatedMembers() 
 
 Dim pvtTable As PivotTable 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 
 ' Determine if calculated members can be viewed. 
 If pvtTable.ViewCalculatedMembers = True Then 
 MsgBox "Calculated members can be viewed." 
 Else 
 MsgBox "Calculated members cannot be viewed." 
 End If 
 
End Sub
```


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)

