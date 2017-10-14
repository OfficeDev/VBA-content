---
title: Workbook.RevisionNumber Property (Excel)
keywords: vbaxl10.chm199139
f1_keywords:
- vbaxl10.chm199139
ms.prod: excel
api_name:
- Excel.Workbook.RevisionNumber
ms.assetid: 7ea9fde5-eb89-a9b0-b637-980f1533d8ec
ms.date: 06/08/2017
---


# Workbook.RevisionNumber Property (Excel)

Returns the number of times the workbook has been saved while open as a shared list. If the workbook is open in exclusive mode, this property returns 0 (zero). Read-only  **Long** .


## Syntax

 _expression_ . **RevisionNumber**

 _expression_ A variable that represents a **Workbook** object.


## Remarks

The  **RevisionNumber** property is updated only when the local copy of the workbook is saved, not when remote copies are saved.


## Example

This example uses the revision number to determine whether the active workbook is open in exclusive mode. If it is, the example saves the workbook as a shared list.


```vb
If ActiveWorkbook.RevisionNumber = 0 Then 
 ActiveWorkbook.SaveAs _ 
 filename:=ActiveWorkbook.FullName, _ 
 accessMode:=xlShared, _ 
 conflictResolution:= _ 
 xlOtherSessionChanges 
End If
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

