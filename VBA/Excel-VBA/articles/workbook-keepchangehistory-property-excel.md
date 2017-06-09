---
title: Workbook.KeepChangeHistory Property (Excel)
keywords: vbaxl10.chm199174
f1_keywords:
- vbaxl10.chm199174
ms.prod: excel
api_name:
- Excel.Workbook.KeepChangeHistory
ms.assetid: 3dbc322e-2b93-ae3d-cb9e-64c657fc0f80
ms.date: 06/08/2017
---


# Workbook.KeepChangeHistory Property (Excel)

 **True** if change tracking is enabled for the shared workbook. Read/write **Boolean** .


## Syntax

 _expression_ . **KeepChangeHistory**

 _expression_ A variable that represents a **Workbook** object.


## Example

This example sets the number of days shown in the change history for the active workbook if change tracking is enabled.


```vb
With ActiveWorkbook 
 If .KeepChangeHistory Then 
 .ChangeHistoryDuration = 7 
 End If 
End With
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

