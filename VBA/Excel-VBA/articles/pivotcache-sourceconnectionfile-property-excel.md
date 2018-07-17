---
title: PivotCache.SourceConnectionFile Property (Excel)
keywords: vbaxl10.chm227103
f1_keywords:
- vbaxl10.chm227103
ms.prod: excel
api_name:
- Excel.PivotCache.SourceConnectionFile
ms.assetid: 87755bde-3c43-3520-24f7-2c778a225b18
ms.date: 06/08/2017
---


# PivotCache.SourceConnectionFile Property (Excel)

Returns or sets a  **String** indicating the Microsoft Office Data Connection file or similar file that was used to create the PivotTable. Read/write.


## Syntax

 _expression_ . **SourceConnectionFile**

 _expression_ A variable that represents a **PivotCache** object.


## Example

This example determines if a connection exists for the PivotTable cache and, if there is a connection, displays the file name. If no connection exists, the code handles the run-time error and notifies the user. This example assumes a PivotTable exists on the active worksheet.


```vb
Sub CheckSourceConnection() 
 
 Dim pvtCache As PivotCache 
 
 Set pvtCache = Application.ActiveWorkbook.PivotCaches.Item(1) 
 
 On Error GoTo No_Connection 
 
 MsgBox "The source connection is: " &; pvtCache.SourceConnectionFile 
 Exit Sub 
 
No_Connection: 
 MsgBox "PivotCache source can not be determined." 
 
End Sub
```


## See also


#### Concepts


[PivotCache Object](pivotcache-object-excel.md)

