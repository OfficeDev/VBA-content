---
title: Range.PageBreak Property (Excel)
keywords: vbaxl10.chm144172
f1_keywords:
- vbaxl10.chm144172
ms.prod: excel
api_name:
- Excel.Range.PageBreak
ms.assetid: 0bec0bba-c2c3-33cd-b39e-55971177c2c8
ms.date: 06/08/2017
---


# Range.PageBreak Property (Excel)

Returns or sets the location of a page break. Can be one of the following  **[XlPageBreak](xlpagebreak-enumeration-excel.md)** constants: **xlPageBreakAutomatic** , **xlPageBreakManual** , or **xlPageBreakNone** . Read/write **Long** .


## Syntax

 _expression_ . **PageBreak**

 _expression_ A variable that represents a **Range** object.


## Remarks

This property can return the location of either automatic or manual page breaks, but it can only set the location of manual breaks (it can only be set to  **xlPageBreakManual** or **xlPageBreakNone** ).

To remove all manual page breaks on a worksheet, set  `Cells.PageBreak` to **xlPageBreakNone** .


## Example

This example sets a manual page break above row 25 on Sheet1.


```vb
Worksheets("Sheet1").Rows(25).PageBreak = xlPageBreakManual
```

This example sets a manual page break to the left of column J on Sheet1.




```vb
Worksheets("Sheet1").Columns("J").PageBreak = xlPageBreakManual
```

This example deletes the two page breaks that were set in the preceding examples.




```vb
Worksheets("Sheet1").Rows(25).PageBreak = xlPageBreakNone 
Worksheets("Sheet1").Columns("J").PageBreak = xlNone
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

