---
title: Protection.AllowUsingPivotTables Property (Excel)
keywords: vbaxl10.chm719083
f1_keywords:
- vbaxl10.chm719083
ms.prod: excel
api_name:
- Excel.Protection.AllowUsingPivotTables
ms.assetid: 42968839-1d82-3c0e-172b-1389c772f9a1
ms.date: 06/08/2017
---


# Protection.AllowUsingPivotTables Property (Excel)

Returns  **True** if the user is allowed to manipulate pivot tables on a protected worksheet. Read-only **Boolean** .


## Syntax

 _expression_ . **AllowUsingPivotTables**

 _expression_ A variable that represents a **Protection** object.


## Remarks

The  **AllowUsingPivotTables** property applies to non-OLAP source data.

The  **AllowUsingPivotTables** property can be set by using the **[Protect](worksheet-protect-method-excel.md)** method arguments.


## Example

This example allows the user to access the PivotTable report and notifies the user. It assumes a non-OLAP Pivot Table report exists on the active worksheet.


```vb
Sub ProtectionOptions() 
 
 ActiveSheet.Unprotect 
 
 ' Allow pivot tables to be manipulated on a protected worksheet. 
 If ActiveSheet.Protection.Allow UsingPivotTables = False Then 
 ActiveSheet.Protect AllowUsingPivotTables:=True 
 End If 
 
 MsgBox "Pivot tables can be manipulated on the protected worksheet." 
 
End Sub
```


## See also


#### Concepts


[Protection Object](protection-object-excel.md)

