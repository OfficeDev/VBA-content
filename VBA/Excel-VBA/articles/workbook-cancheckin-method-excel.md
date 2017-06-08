---
title: Workbook.CanCheckIn Method (Excel)
keywords: vbaxl10.chm199205
f1_keywords:
- vbaxl10.chm199205
ms.prod: excel
api_name:
- Excel.Workbook.CanCheckIn
ms.assetid: 17f7cbdd-0ce0-8e3a-46f3-cb6dafaaa40a
ms.date: 06/08/2017
---


# Workbook.CanCheckIn Method (Excel)

 **True** if Microsoft Excel can check in a specified workbook to a server. Read/write **Boolean** .


## Syntax

 _expression_ . **CanCheckIn**

 _expression_ A variable that represents a **Workbook** object.


### Return Value

Boolean


## Example

This example checks the server to see if the specified workbook can be checked in. If it can be, it saves and closes the workbook and checks it back into the server.


```vb
Sub CheckInOut(strWkbCheckIn As String) 
 
 ' Determine if workbook can be checked in. 
 If Workbooks(strWkbCheckIn).CanCheckIn = True Then 
 Workbooks(strWkbCheckIn).CheckIn 
 MsgBox strWkbCheckIn &; " has been checked in." 
 Else 
 MsgBox "This file cannot be checked in " &; _ 
 "at this time. Please try again later." 
 End If 
 
End Sub
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

