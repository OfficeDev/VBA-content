---
title: Workbook.ChangeHistoryDuration Property (Excel)
keywords: vbaxl10.chm199080
f1_keywords:
- vbaxl10.chm199080
ms.prod: excel
api_name:
- Excel.Workbook.ChangeHistoryDuration
ms.assetid: 5ebc3cc5-dffa-60cf-08cb-b2f84424c4b4
ms.date: 06/08/2017
---


# Workbook.ChangeHistoryDuration Property (Excel)

Returns or sets the number of days shown in the shared workbook's change history. Read/write  **Long** .


## Syntax

 _expression_ . **ChangeHistoryDuration**

 _expression_ A variable that represents a **Workbook** object.


## Remarks

Any changes in the change history older than the setting for this property are removed when the workbook is closed.


## Example

This example sets the number of days shown in the change history for the active workbook if change tracking is enabled. Any changes in the change history older than the setting for this property are removed when the workbook is closed.


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

