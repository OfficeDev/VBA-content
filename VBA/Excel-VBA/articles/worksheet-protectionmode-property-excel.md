---
title: Worksheet.ProtectionMode Property (Excel)
keywords: vbaxl10.chm174092
f1_keywords:
- vbaxl10.chm174092
ms.prod: excel
api_name:
- Excel.Worksheet.ProtectionMode
ms.assetid: 465e2405-c9f3-83ac-f68d-ff9172375e1f
ms.date: 06/08/2017
---


# Worksheet.ProtectionMode Property (Excel)

 **True** if user-interface-only protection is turned on. To turn on user interface protection, use the **[Protect](worksheet-protect-method-excel.md)** method with the _UserInterfaceOnly_ argument set to **True** . Read-only **Boolean** .


## Syntax

 _expression_ . **ProtectionMode**

 _expression_ A variable that represents a **Worksheet** object.


## Example

This example displays the status of the  **ProtectionMode** property.


```vb
MsgBox ActiveSheet.ProtectionMode
```


## See also


#### Concepts


[Worksheet Object](worksheet-object-excel.md)

