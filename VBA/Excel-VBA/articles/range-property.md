---
title: Range Property
keywords: vbagr10.chm65733
f1_keywords:
- vbagr10.chm65733
ms.prod: excel
api_name:
- Excel.Range
ms.assetid: 760f463d-3af3-515d-5da4-54f799fcfe0b
ms.date: 06/08/2017
---


# Range Property

Returns a Range object that represents the specified cell or range of cells. Read-only Range object.

 _expression_. **Range( _Range1_**,  **_Range2_)**

 _expression_ Required. An expression that returns a **DataSheet** object.

 **Range 1** Required for a single cell. The name of the specified range. This must be an A1-style reference in the language the macro is written in. It can include the range operator (a colon), the intersection operator (a space), or the union operator (a comma). It can also include dollar signs, but they're ignored.
OR
 **Range1**,  **_Range2_** Required for a range of cells. The cells in the upper-left and lower-right corners of the specified range. Each argument can be a **Range** object that contains a single cell (or an entire column or entire row), or the argument can be a string that names a single cell in the language the macro is written in.

## Remarks

On the datasheet, the first column heading (starting on the left) is A, followed by B, C, D, and so on. The first row heading (starting at the top) is 1, followed by 2, 3, 4, and so on. Neither the leftmost column nor the top row has a heading. In other words, column A is actually the second column from the left; likewise, row 1 is the second row from the top. The leftmost column and the top row, which are commonly used for legend text or axis labels, are referred to as column 0 (zero) and row 0 (zero). Thus, the following example inserts the text "Annual Sales" in the top cell in column A (the second column).


```
myChart.Application.DataSheet.Range("A0").Value = "Annual Sales"
```

And the following example inserts the text "District 1" in the leftmost cell in row 2 (the third row).




```
myChart.Application.DataSheet.Range("02").Value = "District 1"
```


## Example

This example sets the value of cell A1 on the datasheet to 3.14159.


```
myChart.DataSheet.Range("A1").Value = 3.14159
```

This example loops on cells A1:C3 on the datasheet. If one of the cells has a value less than 0.001, the example replaces that value with 0 (zero).




```vb
With myChart.Application.DataSheet 
 For Each c in .Range("A1:C3") 
 If c.Value < .001 Then 
 c.Value = 0 
 End If 
 Next c 
End With
```


