---
title: Application.MathCoprocessorAvailable Property (Excel)
keywords: vbaxl10.chm133161
f1_keywords:
- vbaxl10.chm133161
ms.prod: excel
api_name:
- Excel.Application.MathCoprocessorAvailable
ms.assetid: 9424d6e1-f6f7-cc1b-7d20-987c8ed5e5a2
ms.date: 06/08/2017
---


# Application.MathCoprocessorAvailable Property (Excel)

 **True** if a math coprocessor is available. Read-only **Boolean** .


## Syntax

 _expression_ . **MathCoprocessorAvailable**

 _expression_ A variable that represents an **Application** object.


## Example

This example displays a message box if a math coprocessor isn't available.


```vb
If Not Application.MathCoprocessorAvailable Then 
 MsgBox "This macro requires a math coprocessor" 
End If
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

