---
title: Worksheet.Protection Property (Excel)
keywords: vbaxl10.chm175154
f1_keywords:
- vbaxl10.chm175154
ms.prod: excel
api_name:
- Excel.Worksheet.Protection
ms.assetid: 46bf2025-46cf-81ae-4864-2d6266dab173
ms.date: 06/08/2017
---


# Worksheet.Protection Property (Excel)

Returns a  **[Protection](protection-object-excel.md)** object that represents the protection options of the worksheet.


## Syntax

 _expression_ . **Protection**

 _expression_ A variable that represents a **Worksheet** object.


## Example

This example protects the active worksheet and then determines if columns can be inserted on the protected worksheet, notifying the user of this status.


```vb
Sub CheckProtection() 
 
 ActiveSheet.Protect 
 
 ' Check the ability to insert columns on a protected sheet. 
 ' Notify the user of this status. 
 If ActiveSheet.Protection.AllowInsertingColumns = True Then 
 MsgBox "The insertion of columns is allowed on this protected worksheet." 
 Else 
 MsgBox "The insertion of columns is not allowed on this protected worksheet." 
 End If 
 
End Sub
```


## See also


#### Concepts


[Worksheet Object](worksheet-object-excel.md)

