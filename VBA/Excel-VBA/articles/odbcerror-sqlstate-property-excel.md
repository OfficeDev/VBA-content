---
title: ODBCError.SqlState Property (Excel)
keywords: vbaxl10.chm527073
f1_keywords:
- vbaxl10.chm527073
ms.prod: excel
api_name:
- Excel.ODBCError.SqlState
ms.assetid: 772a4e82-e661-5568-5fea-49a2925cb156
ms.date: 06/08/2017
---


# ODBCError.SqlState Property (Excel)

Returns the SQL state error. Read-only  **String** .


## Syntax

 _expression_ . **SqlState**

 _expression_ A variable that represents an **ODBCError** object.


## Remarks

For an explanation of the specific error, see you SQL documentation.


## Example

This example refreshes query table one and displays any ODBC errors that occur.


```vb
With Worksheets(1).QueryTables(1) 
 .Refresh 
 Set errs = Application.ODBCErrors 
 If errs.Count > 0 Then 
 Set r = .Destination.Cells(1) 
 r.Value = "The following errors occurred:" 
 c = 0 
 For Each er In errs 
 c = c + 1 
 r.offset(c, 0).value = er.ErrorString 
 r.offset(c, 1).value = er.SqlState 
 Next 
 Else 
 MsgBox "Query complete: all records returned." 
 End If 
End With
```


## See also


#### Concepts


[ODBCError Object](odbcerror-object-excel.md)

