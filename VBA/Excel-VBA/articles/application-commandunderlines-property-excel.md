---
title: Application.CommandUnderlines Property (Excel)
keywords: vbaxl10.chm133095
f1_keywords:
- vbaxl10.chm133095
ms.prod: excel
api_name:
- Excel.Application.CommandUnderlines
ms.assetid: 07d3ea82-6ef4-db6f-f3cf-bef992664408
ms.date: 06/08/2017
---


# Application.CommandUnderlines Property (Excel)

Returns or sets the state of the command underlines in Microsoft Excel for the Macintosh. Can be one of the constants of  **[XlCommandUnderlines](xlcommandunderlines-enumeration-excel.md)** . Read/write **Long** .


## Syntax

 _expression_ . **CommandUnderlines**

 _expression_ A variable that represents an **Application** object.


## Remarks

In Microsoft Excel for Windows, reading this property always returns  **xlCommandUnderlinesOn** , and setting this property to anything other than **xlCommandUnderlinesOn** is an error.


## Example

This example turns off command underlines in Microsoft Excel for the Macintosh.


```vb
Application.CommandUnderlines = xlCommandUnderlinesOff
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

