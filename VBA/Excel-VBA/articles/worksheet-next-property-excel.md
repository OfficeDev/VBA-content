---
title: Worksheet.Next Property (Excel)
keywords: vbaxl10.chm174081
f1_keywords:
- vbaxl10.chm174081
ms.prod: excel
api_name:
- Excel.Worksheet.Next
ms.assetid: 971d5df0-ba23-ac67-7862-67586452e992
ms.date: 06/08/2017
---


# Worksheet.Next Property (Excel)

Returns a  **[Worksheet](worksheet-object-excel.md)** object that represents the next sheet.


## Syntax

 _expression_ . **Next**

 _expression_ A variable that represents a **Worksheet** object.


## Remarks

If the object is a range, this property emulates the TAB key, although the property returns the next cell without selecting it.

On a protected sheet, this property returns the next unlocked cell. On an unprotected sheet, this property always returns the cell immediately to the right of the specified cell.


## Example

This example selects the next unlocked cell on Sheet1. If Sheet1 is unprotected, this is the cell immediately to the right of the active cell.


```vb
Worksheets("Sheet1").Activate 
ActiveCell.Next.Select 

```


## See also


#### Concepts


[Worksheet Object](worksheet-object-excel.md)

