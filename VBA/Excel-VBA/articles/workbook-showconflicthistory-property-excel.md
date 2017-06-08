---
title: Workbook.ShowConflictHistory Property (Excel)
keywords: vbaxl10.chm199153
f1_keywords:
- vbaxl10.chm199153
ms.prod: excel
api_name:
- Excel.Workbook.ShowConflictHistory
ms.assetid: d8588b9e-3e4b-6224-aaa7-ce0b63ff0607
ms.date: 06/08/2017
---


# Workbook.ShowConflictHistory Property (Excel)

 **True** if the Conflict History worksheet is visible in the workbook that's open as a shared list. Read/write **Boolean** .


## Syntax

 _expression_ . **ShowConflictHistory**

 _expression_ A variable that represents a **Workbook** object.


## Remarks

If the specified workbook isn't open as a shared list, this property fails. To determine whether a workbook is open as a shared list, use the  **MultiUserEditing** property.


## Example

This example determines whether the active workbook is open as a shared list. If it is, the example displays the Conflict History worksheet.


```vb
If ActiveWorkbook.MultiUserEditing Then 
 ActiveWorkbook.ShowConflictHistory = True 
End If
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

