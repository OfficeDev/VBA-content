---
title: Protection.AllowDeletingColumns Property (Excel)
keywords: vbaxl10.chm719079
f1_keywords:
- vbaxl10.chm719079
ms.prod: excel
api_name:
- Excel.Protection.AllowDeletingColumns
ms.assetid: 602e0599-f444-0e81-9d9c-70f1f8093a29
ms.date: 06/08/2017
---


# Protection.AllowDeletingColumns Property (Excel)

Returns  **True** if the deletion of columns is allowed on a protected worksheet. Read-only **Boolean** .


## Syntax

 _expression_ . **AllowDeletingColumns**

 _expression_ A variable that represents a **Protection** object.


## Remarks

The  **AllowDeletingColumns** property can be set by using the **[Protect](worksheet-protect-method-excel.md)** method arguments.

The columns containing the cells to be deleted must be unlocked when the sheet is protected.


## Example

This example unlocks column A then allows the user to delete column A on the protected worksheet and notifies the user.


```vb
Sub ProtectionOptions() 
 
 ActiveSheet.Unprotect 
 
 'Unlock column A. 
 Columns("A:A").Locked = False 
 
 ' Allow column A to be deleted on a protected worksheet. 
 If ActiveSheet.Protection.AllowDeletingColumns = False Then 
 ActiveSheet.Protect AllowDeletingColumns:=True 
 End If 
 
 MsgBox "Column A can be deleted on this protected worksheet." 
 
End Sub
```


## See also


#### Concepts


[Protection Object](protection-object-excel.md)

