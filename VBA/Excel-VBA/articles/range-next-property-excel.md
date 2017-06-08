---
title: Range.Next Property (Excel)
keywords: vbaxl10.chm144165
f1_keywords:
- vbaxl10.chm144165
ms.prod: excel
api_name:
- Excel.Range.Next
ms.assetid: 10712827-9abd-6b8a-49e5-65e3554fcd87
ms.date: 06/08/2017
---


# Range.Next Property (Excel)

Returns a  **[Range](range-object-excel.md)** object that represents the next cell.


## Syntax

 _expression_ . **Next**

 _expression_ A variable that represents a **Range** object.


## Remarks

If the object is a range, this property emulates the TAB key, although the property returns the next cell without selecting it.

On a protected sheet, this property returns the next unlocked cell. On an unprotected sheet, this property always returns the cell immediately to the right of the specified cell.


## See also


#### Concepts


[Range Object](range-object-excel.md)

