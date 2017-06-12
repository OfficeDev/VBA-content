---
title: Workbook.ConflictResolution Property (Excel)
keywords: vbaxl10.chm199091
f1_keywords:
- vbaxl10.chm199091
ms.prod: excel
api_name:
- Excel.Workbook.ConflictResolution
ms.assetid: 5142c848-0731-14d9-5913-bbaa67bf308f
ms.date: 06/08/2017
---


# Workbook.ConflictResolution Property (Excel)

Returns or sets the way conflicts are to be resolved whenever a shared workbook is updated. Read/write  **[XlSaveConflictResolution](xlsaveconflictresolution-enumeration-excel.md)** .


## Syntax

 _expression_ . **ConflictResolution**

 _expression_ A variable that represents a **Workbook** object.


## Remarks





| **XlSaveConflictResolution** can be one of these **XlSaveConflictResolution** constants.|
| **xlLocalSessionChanges** . The local user's changes are always accepted.|
| **xlOtherSessionChanges** . The local user's changes are always rejected.|
| **xlUserResolution** . A dialog box asks the user to resolve the conflict.|

## Example

This example causes the local user's changes to be accepted whenever there's a conflict in the shared workbook.


```vb
ActiveWorkbook.ConflictResolution = xlLocalSessionChanges 

```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

