---
title: PivotTable.EnableDataValueEditing Property (Excel)
keywords: vbaxl10.chm235141
f1_keywords:
- vbaxl10.chm235141
ms.prod: excel
api_name:
- Excel.PivotTable.EnableDataValueEditing
ms.assetid: 57b4ed51-46d5-0d9f-d947-cdc45e523095
ms.date: 06/08/2017
---


# PivotTable.EnableDataValueEditing Property (Excel)

 **True** to disable the alert for when the user overwrites values in the data area of the PivotTable. **True** also allows the user to change data values that previously could not be changed. The default value is **False** . Read/write **Boolean** .


## Syntax

 _expression_ . **EnableDataValueEditing**

 _expression_ A variable that represents a **PivotTable** object.


## Remarks

Any editing performed on data values is lost upon refresh.


## Example

This example determines the alert setting for overwriting values in the data area and notifies the user. The example assumes a PivotTable exists on the active worksheet.


```vb
Sub CheckAlertSetting() 
 
 Dim pvtTable As PivotTable 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 
 ' Determine alert setting. 
 If pvtTable.EnableDataValueEditing = False Then 
 MsgBox "Alert is enabled." 
 Else 
 MsgBox "Alert is disabled." 
 End If 
 
End Sub
```


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)

