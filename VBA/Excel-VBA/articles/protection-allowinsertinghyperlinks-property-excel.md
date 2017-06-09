---
title: Protection.AllowInsertingHyperlinks Property (Excel)
keywords: vbaxl10.chm719078
f1_keywords:
- vbaxl10.chm719078
ms.prod: excel
api_name:
- Excel.Protection.AllowInsertingHyperlinks
ms.assetid: ef334ce3-a8d3-d9db-e48b-739f150cfb98
ms.date: 06/08/2017
---


# Protection.AllowInsertingHyperlinks Property (Excel)

Returns  **True** if the insertion of hyperlinks is allowed on a protected worksheet. Read-only **Boolean** .


## Syntax

 _expression_ . **AllowInsertingHyperlinks**

 _expression_ A variable that represents a **Protection** object.


## Remarks

Hyperlinks can only be inserted in unlocked or unprotected cells on a protected worksheet.

The  **AllowInsertingHyperlinks** property can be set by using the **[Protect](worksheet-protect-method-excel.md)** method arguments.


## Example

This example allows the user to insert a hyperlink in cell A1 on the protected worksheet and notifies the user.


```vb
Sub ProtectionOptions() 
 
 ActiveSheet.Unprotect 
 
 ' Unlock cell A1. 
 Range("A1").Locked = False 
 
 ' Allow hyperlinks to be inserted on a protected worksheet. 
 If ActiveSheet.Protection.AllowInsertingHyperlinks = False Then 
 ActiveSheet.Protect AllowInsertingHyperlinks:=True 
 End If 
 
 MsgBox "Hyperlinks can be inserted on this protected worksheet." 
 
End Sub
```


## See also


#### Concepts


[Protection Object](protection-object-excel.md)

