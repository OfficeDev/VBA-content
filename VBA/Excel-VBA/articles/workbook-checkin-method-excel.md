---
title: Workbook.CheckIn Method (Excel)
keywords: vbaxl10.chm199204
f1_keywords:
- vbaxl10.chm199204
ms.prod: excel
api_name:
- Excel.Workbook.CheckIn
ms.assetid: f9750086-aaa6-3b04-6b51-ebcadf6b1911
ms.date: 06/08/2017
---


# Workbook.CheckIn Method (Excel)

Returns a workbook from a local computer to a server, and sets the local workbook to read-only so that it cannot be edited locally. Calling this method will also close the workbook.


## Syntax

 _expression_ . **CheckIn**( **_SaveChanges_** , **_Comments_** , **_MakePublic_** )

 _expression_ A variable that represents a **Workbook** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SaveChanges_|Optional| **Variant**| **True** saves changes and checks in the document. **False** returns the document to a checked-in status without saving revision.|
| _Comments_|Optional| **Variant**|Allows the user to enter check-in comments for the revision of the workbook being checked in (applies only if  _SaveChanges_ equals **True** ).|
| _MakePublic_|Optional| **Variant**| **True** allows the user to publish the workbook after it has been checked in. This submits the workbook for the approval process, which can eventually result in a version of the workbook being published to users with read-only rights to the workbook (applies only if _SaveChanges_ equals **True** ).|

## Example

This example checks the server to see if the specified workbook can be checked in. If it can, the code saves and closes the workbook and checks it back in to the server.


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

