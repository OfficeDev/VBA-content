---
title: Chart.Next Property (Excel)
keywords: vbaxl10.chm148081
f1_keywords:
- vbaxl10.chm148081
ms.prod: excel
api_name:
- Excel.Chart.Next
ms.assetid: a0e53eba-c9e9-7997-4765-90debeb8ae5d
ms.date: 06/08/2017
---


# Chart.Next Property (Excel)

Returns a  **[Worksheet](worksheet-object-excel.md)** object that represents the next sheet.


## Syntax

 _expression_ . **Next**

 _expression_ A variable that represents a **Chart** object.


## Remarks

If the object is a range, this property emulates the TAB key, although the property returns the next cell without selecting it.

On a protected sheet, this property returns the next unlocked cell. On an unprotected sheet, this property always returns the cell immediately to the right of the specified cell.


## See also


#### Concepts


[Chart Object](chart-object-excel.md)

