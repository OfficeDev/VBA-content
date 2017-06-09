---
title: ColumnWidth Property
keywords: vbagr10.chm65778
f1_keywords:
- vbagr10.chm65778
ms.prod: excel
api_name:
- Excel.ColumnWidth
ms.assetid: fffb3493-4b40-7a0b-f3ad-d191baebb87f
ms.date: 06/08/2017
---


# ColumnWidth Property

Returns or sets the width of all columns in the specified range. Read/write Variant.

 _expression_. **ColumnWidth**

 _expression_ Required. An expression that returns one of the objects in the Applies To list.


## Remarks

One unit of column width is equal to the width of one character in the Normal style. For proportional fonts, the width of the character 0 (zero) is used.

If all columns in the range have the same width, the  **ColumnWidth** property returns the width. If columns in the range have different widths, this property returns **Null**.


## Example

This example doubles the width of column A on the datasheet.


```vb
With myChart.Application.DataSheet.Columns("A") 
 .ColumnWidth = .ColumnWidth * 2 
End With
```


