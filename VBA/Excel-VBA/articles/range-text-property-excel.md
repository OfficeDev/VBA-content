---
title: Range.Text Property (Excel)
keywords: vbaxl10.chm144209
f1_keywords:
- vbaxl10.chm144209
ms.prod: excel
api_name:
- Excel.Range.Text
ms.assetid: e38c15b1-5941-0a28-1acf-328bc214a2e0
ms.date: 06/08/2017
---


# Range.Text Property (Excel)

Returns the formatted text for the specified object. Read-only  **String**.


## Syntax

 _expression_ . **Text**

 _expression_ A variable that represents a **Range** object.


## Remarks

The **Text** property is most often used for a **Range** of one cell. If the **Range** includes more than one cell, the **Text** property returns **Null**, except when all the cells in the **Range** have identical contents and formats.


## Example

This example illustrates the difference between the  **Text** and **[Value](range-value-property-excel.md)** properties of cells that contain formatted numbers.


```vb
Dim c As Range
Set c = Worksheets("Sheet1").Range("A1")
c.Value = 1198.3
c.NumberFormat = "$#,##0_);($#,##0)"
MsgBox c.Value & " is the value." 'Returns "1198.3 is the value."
MsgBox c.Text & " is the text."   'Returns "$1,198 is the text."
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

[Range.Value Property](range-value-property-excel.md)
