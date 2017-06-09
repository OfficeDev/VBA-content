---
title: Workbooks.CanCheckOut Method (Excel)
keywords: vbaxl10.chm203086
f1_keywords:
- vbaxl10.chm203086
ms.prod: excel
api_name:
- Excel.Workbooks.CanCheckOut
ms.assetid: 06259bbd-2e55-3fbe-d08c-094985cb9fff
ms.date: 06/08/2017
---


# Workbooks.CanCheckOut Method (Excel)

 **True** if Microsoft Excel can check out a specified workbook from a server. Read/write **Boolean** .


## Syntax

 _expression_ . **CanCheckOut**( **_Filename_** )

 _expression_ A variable that represents a **Workbooks** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Filename_|Required| **String**|The name of the file to check out.|

### Return Value

Boolean


## Example

This example verifies that a workbook is not checked out by another user and can be checked out. If the workbook can be checked out, it copies the workbook to the local computer for editing.


```vb
Sub UseCanCheckOut(docCheckOut As String) 
 
 ' Determine if workbook can be checked out. 
 If Workbooks.CanCheckOut(Filename:=docCheckOut) = True Then 
 Workbooks.CheckOut (Filename:=docCheckOut) 
 Else 
 MsgBox "You are unable to check out this document at this time." 
 End If 
 
End Sub
```


## See also


#### Concepts


[Workbooks Object](workbooks-object-excel.md)

