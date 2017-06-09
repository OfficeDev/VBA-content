---
title: Application.StandardFontSize Property (Excel)
keywords: vbaxl10.chm133211
f1_keywords:
- vbaxl10.chm133211
ms.prod: excel
api_name:
- Excel.Application.StandardFontSize
ms.assetid: 368ae001-7471-d104-573a-fc97d761f75e
ms.date: 06/08/2017
---


# Application.StandardFontSize Property (Excel)

Returns or sets the standard font size, in points. Read/write  **Long** .


## Syntax

 _expression_ . **StandardFontSize**

 _expression_ A variable that represents an **Application** object.


## Remarks

If you change the standard font size by using this property, the change doesn't take effect until you restart Microsoft Excel.


## Example

This example sets the standard font size to 12 points.


```vb
Application.StandardFontSize = 12
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

