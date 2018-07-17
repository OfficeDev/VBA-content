---
title: Protection.AllowSorting Property (Excel)
keywords: vbaxl10.chm719081
f1_keywords:
- vbaxl10.chm719081
ms.prod: excel
api_name:
- Excel.Protection.AllowSorting
ms.assetid: cffdb62d-2fbb-111a-ed06-e295b722ee75
ms.date: 06/08/2017
---


# Protection.AllowSorting Property (Excel)

Returns  **True** if the sorting option is allowed on a protected worksheet. Read-only **Boolean** .


## Syntax

 _expression_ . **AllowSorting**

 _expression_ A variable that represents a **Protection** object.


## Remarks

Sorting can only be performed on unlocked or unprotected cells in a protected worksheet.

The  **AllowSorting** property can be set by using the **[Protect](worksheet-protect-method-excel.md)** method arguments.


## Example

This example allows the user to sort unlocked or unprotected cells on the protected worksheet and notifies the user.


```vb
Sub ProtectionOptions() 
 
 ActiveSheet.Unprotect 
 
 ' Unlock cells A1 through B5. 
 Range("A1:B5").Locked = False 
 
 ' Allow sorting to be performed on the protected worksheet. 
 If ActiveSheet.Protection.AllowSorting = False Then 
 ActiveSheet.Protect AllowSorting:=True 
 End If 
 
 MsgBox "For cells A1 through B5, sorting can be performed on the protected worksheet." 
 
End Sub
```


## See also


#### Concepts


[Protection Object](protection-object-excel.md)

