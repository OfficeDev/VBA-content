---
title: Workbook.MultiUserEditing Property (Excel)
keywords: vbaxl10.chm199113
f1_keywords:
- vbaxl10.chm199113
ms.prod: excel
api_name:
- Excel.Workbook.MultiUserEditing
ms.assetid: dc721463-ec34-8c52-6701-51c406beed23
ms.date: 06/08/2017
---


# Workbook.MultiUserEditing Property (Excel)

 **True** if the workbook is open as a shared list. Read-only **Boolean** .


## Syntax

 _expression_ . **MultiUserEditing**

 _expression_ A variable that represents a **Workbook** object.


## Remarks

To save a workbook as a shared list, use the  **SaveAs** method. To switch the workbook from shared mode to exclusive mode, use the **ExclusiveAccess** method.


## Example

This example determines whether the active workbook is open in exclusive mode. If it is, the example saves the workbook as a shared list.


```vb
If Not ActiveWorkbook.MultiUserEditing Then 
 ActiveWorkbook.SaveAs fileName:=ActiveWorkbook.FullName, _ 
 accessMode:=xlShared 
End If
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

