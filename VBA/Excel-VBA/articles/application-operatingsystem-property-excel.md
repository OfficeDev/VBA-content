---
title: Application.OperatingSystem Property (Excel)
keywords: vbaxl10.chm133187
f1_keywords:
- vbaxl10.chm133187
ms.prod: excel
api_name:
- Excel.Application.OperatingSystem
ms.assetid: a36c5080-1d7e-a941-1bad-94f92522c7cf
ms.date: 06/08/2017
---


# Application.OperatingSystem Property (Excel)

Returns the name and version number of the current operating system — for example, "Windows (32-bit) 4.00" or "Macintosh 7.00".

Return e.g. "Windows (32-bit) NT 6.02" with Win8.1 (=6.02, **64bit**) and Excel 2013 (15.0.4631.1000, 32bit)

E.g. "Windows (64-bit) NT :.00" with Win10 (64bit) and Excel 2016 (16.0.6326.1010, 64bit)

Read-only  **String** .


## Syntax

 _expression_ . **OperatingSystem**

 _expression_ A variable that represents an **Application** object.


## Example

This example displays the name of the operating system.


```vb
MsgBox "Microsoft Excel is using " & Application.OperatingSystem
```

## See also


#### Concepts


[Application Object](application-object-excel.md)

