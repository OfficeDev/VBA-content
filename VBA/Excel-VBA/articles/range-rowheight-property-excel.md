---
title: Range.RowHeight Property (Excel)
keywords: vbaxl10.chm144190
f1_keywords:
- vbaxl10.chm144190
ms.prod: excel
api_name:
- Excel.Range.RowHeight
ms.assetid: 103c7209-9a4f-8f9c-7bdc-3013113867a5
ms.date: 06/08/2017
---


# Range.RowHeight Property (Excel)

Returns or sets the height of rows in the range specified, measured in points. Read/write **Variant**.


## Syntax

 _expression_ . **RowHeight**

 _expression_ A variable that represents a **Range** object.


## Remarks

**RowHeight** sets the height for all rows in a range.

Use the  **[Height](range-height-property-excel.md)** property to return the total height of a range of cells.

**Note**  When a range contains rows of different heights, **RowHeight** might return the height of the first row or might return **Null**. 


## Example

This example doubles the height of row one on Sheet1.


```vb
With Worksheets("Sheet1").Rows(1) 
 .RowHeight = .RowHeight * 2 
End With
```


## See also


#### Concepts

[Range Object](range-object-excel.md)

[Range.Height Property](range-height-property-excel.md)

[Slicer.RowHeight Property](slicer-rowheight-property-excel.md)
