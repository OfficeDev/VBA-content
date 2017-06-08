---
title: PivotField.EnableItemSelection Property (Excel)
keywords: vbaxl10.chm240134
f1_keywords:
- vbaxl10.chm240134
ms.prod: excel
api_name:
- Excel.PivotField.EnableItemSelection
ms.assetid: ae55f88a-618f-3063-2019-a993a3146b67
ms.date: 06/08/2017
---


# PivotField.EnableItemSelection Property (Excel)

When set to  **False** , disables the ability to use the field dropdown in the user interface. The default value is **True** . Read/write **Boolean** .


## Syntax

 _expression_ . **EnableItemSelection**

 _expression_ A variable that represents a **PivotField** object.


## Remarks

A run-time error will occur if the OLAP PivotTable field is not the highest level for the hierarchy.


## Example

This example determines the setting for selecting items using the field dropdown and enables the feature, if necessary. The example assumes a PivotTable exists on the active worksheet.


```vb
Sub UseEnableItemSelection() 
 
 Dim pvtTable As PivotTable 
 Dim pvtField As PivotField 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 Set pvtField = pvtTable.RowFields(1) 
 
 ' Determine setting for property and enable if necessary. 
 If pvtField.EnableItemSelection = False Then 
 pvtField.EnableItemSelection = True 
 MsgBox "Item selection enabled for fields." 
 Else 
 MsgBox "Item selection is already enabled for fields." 
 End If 
 
End Sub
```


## See also


#### Concepts


[PivotField Object](pivotfield-object-excel.md)

