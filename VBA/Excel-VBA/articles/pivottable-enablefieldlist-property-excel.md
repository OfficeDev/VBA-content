---
title: PivotTable.EnableFieldList Property (Excel)
keywords: vbaxl10.chm235148
f1_keywords:
- vbaxl10.chm235148
ms.prod: excel
api_name:
- Excel.PivotTable.EnableFieldList
ms.assetid: 3f078d19-d2ec-1c1a-e039-69e8d7e21e95
ms.date: 06/08/2017
---


# PivotTable.EnableFieldList Property (Excel)

 **False** to disable the ability to display the field list for the PivotTable. If the field list was already being displayed it disappears. The default value is **True** . Read/write **Boolean** .


## Syntax

 _expression_ . **EnableFieldList**

 _expression_ A variable that represents a **PivotTable** object.


## Example

This example determines the whether the field list for the PivotTable can be displayed or not and notifies the user. The example assumes that a PivotTable exists on the active worksheet.


```vb
Sub CheckFieldList() 
 
 Dim pvtTable As PivotTable 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 
 ' Determine if field list can be displayed. 
 If pvtTable.EnableFieldList = True Then 
 MsgBox "The field list for the PivotTable can be displayed." 
 Else 
 MsgBox "The field list for the PivotTable cannot be displayed." 
 End If 
 
End Sub
```


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)

