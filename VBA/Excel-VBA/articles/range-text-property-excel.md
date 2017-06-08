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

Returns or sets the text for the specified object. Read-only  **String** .


## Syntax

 _expression_ . **Text**

 _expression_ A variable that represents a **Range** object.


## Example

This example illustrates the difference between the  **Text** and **Value** properties of cells that contain formatted numbers.


```vb
Set c = Worksheets("Sheet1").Range("B14") 
c.Value = 1198.3 
c.NumberFormat = "$#,##0_);($#,##0)" 
MsgBox c.Value 
MsgBox c.Text
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

