---
title: CellFormat.NumberFormat Property (Excel)
keywords: vbaxl10.chm676076
f1_keywords:
- vbaxl10.chm676076
ms.prod: excel
api_name:
- Excel.CellFormat.NumberFormat
ms.assetid: 55133c7e-7d55-a2a9-0a76-9bd630a59cc4
ms.date: 06/08/2017
---


# CellFormat.NumberFormat Property (Excel)

Returns or sets a  **Variant** value that represents the format code for the object.


## Syntax

 _expression_ . **NumberFormat**

 _expression_ A variable that represents a **CellFormat** object.


## Remarks

This property returns  **Null** if all cells in the specified range don't have the same number format.

The format code is the same string as the  **Format Codes** option in the **Format Cells** dialog box. The **Format** function uses different format code strings than do the **NumberFormat** and **[NumberFormatLocal](cellformat-numberformatlocal-property-excel.md)** properties.


## See also


#### Concepts


[CellFormat Object](cellformat-object-excel.md)

