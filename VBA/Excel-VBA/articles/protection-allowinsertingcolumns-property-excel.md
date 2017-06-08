---
title: Protection.AllowInsertingColumns Property (Excel)
keywords: vbaxl10.chm719076
f1_keywords:
- vbaxl10.chm719076
ms.prod: excel
api_name:
- Excel.Protection.AllowInsertingColumns
ms.assetid: 87938c66-e48a-dd1d-934e-08752bbf3e03
ms.date: 06/08/2017
---


# Protection.AllowInsertingColumns Property (Excel)

Returns  **True** if the insertion of columns is allowed on a protected worksheet. Read-only **Boolean** .


## Syntax

 _expression_ . **AllowInsertingColumns**

 _expression_ A variable that represents a **Protection** object.


## Remarks

An inserted column inherits its formatting (by default) from the column to its left, which means that it may have locked cells. In other words, users may not be able to delete columns that they have inserted.

The  **AllowInsertingColumns** property can be set by using the **[Protect](worksheet-protect-method-excel.md)** method arguments.


## Example

This example allows the user to insert columns on the protected worksheet and notifies the user.


```vb
Sub ProtectionOptions() 
 
 ActiveSheet.Unprotect 
 
 ' Allow columns to be inserted on a protected worksheet. 
 If ActiveSheet.Protection.AllowInsertingColumns = False Then 
 ActiveSheet.Protect AllowInsertingColumns:=True 
 End If 
 
 MsgBox "Columns can be inserted on this protected worksheet." 
 
End Sub
```


## See also


#### Concepts


[Protection Object](protection-object-excel.md)

