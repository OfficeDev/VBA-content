---
title: CellFormat.WrapText Property (Excel)
keywords: vbaxl10.chm676084
f1_keywords:
- vbaxl10.chm676084
ms.prod: excel
api_name:
- Excel.CellFormat.WrapText
ms.assetid: 92d7920c-51e2-f949-60ee-d11595c191bb
ms.date: 06/08/2017
---


# CellFormat.WrapText Property (Excel)

Returns or sets a  **Variant** value that indicates if Microsoft Excel wraps the text in the object.


## Syntax

 _expression_ . **WrapText**

 _expression_ A variable that represents a **CellFormat** object.


## Remarks

This property returns  **True** if text is wrapped in all cells within the specified range, **False** if text is not wrapped in all cells within the specified range, or **Null** if the specified range contains some cells that wrap text and other cells that don't.

Microsoft Excel will change the row height of the range, if necessary, to accommodate the text in the range.


## See also


#### Concepts


[CellFormat Object](cellformat-object-excel.md)

