---
title: Workbook.UserStatus Property (Excel)
keywords: vbaxl10.chm199163
f1_keywords:
- vbaxl10.chm199163
ms.prod: excel
api_name:
- Excel.Workbook.UserStatus
ms.assetid: 0df24f8a-b60b-fd8c-3436-903652487a09
ms.date: 06/08/2017
---


# Workbook.UserStatus Property (Excel)

Returns a 1-based, two-dimensional array that provides information about each user who has the workbook open as a shared list. Read-only  **Variant** .


## Syntax

 _expression_ . **UserStatus**

 _expression_ A variable that represents a **Workbook** object.


## Remarks

The first element of the second dimension is the name of the user, the second element is the date and time when the user last opened the workbook, and the third element is a number indicating the type of list (1 indicates exclusive, and 2 indicates shared).

The  **UserStatus** property doesn't return information about users who have the specified workbook open as read-only.


## Example

This example creates a new workbook and inserts into it information about all users who have the active workbook open as a shared list.


```vb
users = ActiveWorkbook.UserStatus 
With Workbooks.Add.Sheets(1) 
 For row = 1 To UBound(users, 1) 
 .Cells(row, 1) = users(row, 1) 
 .Cells(row, 2) = users(row, 2) 
 Select Case users(row, 3) 
 Case 1 
 .Cells(row, 3).Value = "Exclusive" 
 Case 2 
 .Cells(row, 3).Value = "Shared" 
 End Select 
 Next 
End With
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

