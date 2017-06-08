---
title: Protection.AllowDeletingRows Property (Excel)
keywords: vbaxl10.chm719080
f1_keywords:
- vbaxl10.chm719080
ms.prod: excel
api_name:
- Excel.Protection.AllowDeletingRows
ms.assetid: da418f4e-ca3e-b0f2-4b12-fe578b0bf20b
ms.date: 06/08/2017
---


# Protection.AllowDeletingRows Property (Excel)

Returns  **True** if the deletion of rows is allowed on a protected worksheet. Read-only **Boolean** .


## Syntax

 _expression_ . **AllowDeletingRows**

 _expression_ A variable that represents a **Protection** object.


## Remarks

The  **AllowDeletingRows** property can be set by using the **[Protect](worksheet-protect-method-excel.md)** method arguments.

The rows containing the cells to be deleted must be unlocked when the sheet is protected.


## Example

This example unlocks row 1 then allows the user to delete row 1 on the protected worksheet and notifies the user.


```vb
Sub ProtectionOptions() 
 
 ActiveSheet.Unprotect 
 
 'Unlock row 1. 
 Rows("1:1").Locked = False 
 
 ' Allow row 1 to be deleted on a protected worksheet. 
 If ActiveSheet.Protection.AllowDeletingRows = False Then 
 ActiveSheet.Protect AllowDeletingRows:=True 
 End If 
 
 MsgBox "Row 1 can be deleted on this protected worksheet." 
 
End Sub
```


## See also


#### Concepts


[Protection Object](protection-object-excel.md)

