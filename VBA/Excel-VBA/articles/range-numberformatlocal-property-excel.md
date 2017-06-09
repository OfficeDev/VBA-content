---
title: Range.NumberFormatLocal Property (Excel)
keywords: vbaxl10.chm144168
f1_keywords:
- vbaxl10.chm144168
ms.prod: excel
api_name:
- Excel.Range.NumberFormatLocal
ms.assetid: e34e6f52-9279-7961-adfa-4aa84c44937a
ms.date: 06/08/2017
---


# Range.NumberFormatLocal Property (Excel)

Returns or sets a  **Variant** value that represents the format code for the object as a string in the language of the user.


## Syntax

 _expression_ . **NumberFormatLocal**

 _expression_ A variable that represents a **Range** object.


## Remarks

The  **Format** function uses different format code strings than do the **[NumberFormat](range-numberformat-property-excel.md)** and **NumberFormatLocal** properties.


## Example

This example displays the number format for cell A1 on Sheet1 in the language of the user.


```vb
MsgBox "The number format for cell A1 is " &; _ 
 Worksheets("Sheet1").Range("A1").NumberFormatLocal
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

