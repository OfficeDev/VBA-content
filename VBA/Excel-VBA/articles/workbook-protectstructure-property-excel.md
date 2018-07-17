---
title: Workbook.ProtectStructure Property (Excel)
keywords: vbaxl10.chm199131
f1_keywords:
- vbaxl10.chm199131
ms.prod: excel
api_name:
- Excel.Workbook.ProtectStructure
ms.assetid: bf721b60-0ad1-f71c-7ef4-74d2196d320e
ms.date: 06/08/2017
---


# Workbook.ProtectStructure Property (Excel)

 **True** if the order of the sheets in the workbook is protected. Read-only **Boolean** .


## Syntax

 _expression_ . **ProtectStructure**

 _expression_ A variable that represents a **Workbook** object.


## Example

This example displays a message if the order of the sheets in the active workbook is protected.


```vb
If ActiveWorkbook.ProtectStructure = True Then 
 MsgBox "Remember, you cannot delete, add, or change " &; _ 
 Chr(13) &; _ 
 "the location of any sheets in this workbook." 
End If
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

