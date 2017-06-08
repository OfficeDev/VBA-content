---
title: Protection.AllowInsertingRows Property (Excel)
keywords: vbaxl10.chm719077
f1_keywords:
- vbaxl10.chm719077
ms.prod: excel
api_name:
- Excel.Protection.AllowInsertingRows
ms.assetid: 481fb5d0-31c9-9c28-c5a0-3f3abc48ad3a
ms.date: 06/08/2017
---


# Protection.AllowInsertingRows Property (Excel)

Returns  **True** if the insertion of rows is allowed on a protected worksheet. Read-only **Boolean** .


## Syntax

 _expression_ . **AllowInsertingRows**

 _expression_ A variable that represents a **Protection** object.


## Remarks

The  **AllowInsertingRows** property can be set by using the **[Protect](worksheet-protect-method-excel.md)** method arguments.


## Example

This example allows the user to insert rows on the protected worksheet and notifies the user.


```vb
Sub ProtectionOptions() 
 
 ActiveSheet.Unprotect 
 
 ' Allow rows to be inserted on a protected worksheet. 
 If ActiveSheet.Protection.AllowInsertingRows = False Then 
 ActiveSheet.Protect AllowInsertingRows:=True 
 End If 
 
 MsgBox "Rows can be inserted on this protected worksheet." 
 
End Sub
```


## See also


#### Concepts


[Protection Object](protection-object-excel.md)

