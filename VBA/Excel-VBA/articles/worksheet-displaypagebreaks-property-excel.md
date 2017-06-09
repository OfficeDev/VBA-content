---
title: Worksheet.DisplayPageBreaks Property (Excel)
keywords: vbaxl10.chm175138
f1_keywords:
- vbaxl10.chm175138
ms.prod: excel
api_name:
- Excel.Worksheet.DisplayPageBreaks
ms.assetid: 95152278-2618-f200-9933-b6574a49e256
ms.date: 06/08/2017
---


# Worksheet.DisplayPageBreaks Property (Excel)

 **True** if page breaks (both automatic and manual) on the specified worksheet are displayed. Read/write **Boolean** .


## Syntax

 _expression_ . **DisplayPageBreaks**

 _expression_ A variable that represents a **Worksheet** object.


## Remarks

You can't set this property if you don't have a printer installed.


## Example

This example causes Sheet1 to display page breaks.


```vb
Worksheets("Sheet1").DisplayPageBreaks = True
```


## See also


#### Concepts


[Worksheet Object](worksheet-object-excel.md)

