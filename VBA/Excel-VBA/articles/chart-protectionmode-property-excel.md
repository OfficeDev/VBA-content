---
title: Chart.ProtectionMode Property (Excel)
keywords: vbaxl10.chm148092
f1_keywords:
- vbaxl10.chm148092
ms.prod: excel
api_name:
- Excel.Chart.ProtectionMode
ms.assetid: 5a9afe8c-df46-cbfe-d692-d4be8f2e505b
ms.date: 06/08/2017
---


# Chart.ProtectionMode Property (Excel)

 **True** if user-interface-only protection is turned on. To turn on user interface protection, use the **[Protect](chart-protect-method-excel.md)** method with the _UserInterfaceOnly_ argument set to **True** . Read-only **Boolean** .


## Syntax

 _expression_ . **ProtectionMode**

 _expression_ A variable that represents a **Chart** object.


## Example

This example displays the status of the  **ProtectionMode** property.


```vb
MsgBox ActiveSheet.ProtectionMode
```


## See also


#### Concepts


[Chart Object](chart-object-excel.md)

