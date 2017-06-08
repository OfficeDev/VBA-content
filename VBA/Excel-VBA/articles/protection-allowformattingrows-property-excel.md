---
title: Protection.AllowFormattingRows Property (Excel)
keywords: vbaxl10.chm719075
f1_keywords:
- vbaxl10.chm719075
ms.prod: excel
api_name:
- Excel.Protection.AllowFormattingRows
ms.assetid: c58f9511-b6f5-a911-d20d-90dbb46248b7
ms.date: 06/08/2017
---


# Protection.AllowFormattingRows Property (Excel)

Returns  **True** if the formatting of rows is allowed on a protected worksheet. Read-only **Boolean** .


## Syntax

 _expression_ . **AllowFormattingRows**

 _expression_ A variable that represents a **Protection** object.


## Remarks

The  **AllowFormattingRows** property can be set by using the **[Protect](worksheet-protect-method-excel.md)** method arguments.


## Example

This example allows the user to format the rows on the protected worksheet and notifies the user.


```vb
Sub ProtectionOptions() 
 
 ActiveSheet.Unprotect 
 
 ' Allow rows to be formatted on a protected worksheet. 
 If ActiveSheet.Protection.AllowFormattingRows = False Then 
 ActiveSheet.Protect AllowFormattingRows:=True 
 End If 
 
 MsgBox "Rows can be formatted on this protected worksheet." 
 
End Sub
```


## See also


#### Concepts


[Protection Object](protection-object-excel.md)

