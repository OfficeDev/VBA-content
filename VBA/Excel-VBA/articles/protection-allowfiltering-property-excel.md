---
title: Protection.AllowFiltering Property (Excel)
keywords: vbaxl10.chm719082
f1_keywords:
- vbaxl10.chm719082
ms.prod: excel
api_name:
- Excel.Protection.AllowFiltering
ms.assetid: dc0b8ab3-ea28-0692-9474-8f81cc395599
ms.date: 06/08/2017
---


# Protection.AllowFiltering Property (Excel)

Returns  **True** if the user is allowed to make use of an AutoFilter that was created before the sheet was protected. Read-only **Boolean** .


## Syntax

 _expression_ . **AllowFiltering**

 _expression_ A variable that represents a **Protection** object.


## Remarks

The  **AllowFiltering** property can be set by using the **[Protect](worksheet-protect-method-excel.md)** method arguments.

The  **AllowFiltering** property allows the user to change filter criteria on an existing AutoFilter. The user cannot create or remove an AutoFilter on a protected worksheet.

The cells to be filtered must be unlocked when the sheet is protected.


## Example

This example allows the user to filter row 1 on the protected worksheet and notifies the user.


```vb
Sub ProtectionOptions() 
 
 ActiveSheet.Unprotect 
 
 ' Unlock row 1. 
 Rows("1:1").Locked = False 
 
 ' Allow row 1 to be filtered on a protected worksheet. 
 If ActiveSheet.Protection.AllowFiltering = False Then 
 ActiveSheet.Protect AllowFiltering:=True 
 End If 
 
 MsgBox "Row 1 can be filtered on this protected worksheet." 
 
End Sub
```


## See also


#### Concepts


[Protection Object](protection-object-excel.md)

