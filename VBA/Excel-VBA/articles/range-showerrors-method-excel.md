---
title: Range.ShowErrors Method (Excel)
keywords: vbaxl10.chm144197
f1_keywords:
- vbaxl10.chm144197
ms.prod: excel
api_name:
- Excel.Range.ShowErrors
ms.assetid: 02366ef0-b4dc-a10c-e186-d9392a8b656c
ms.date: 06/08/2017
---


# Range.ShowErrors Method (Excel)

Draws tracer arrows through the precedents tree to the cell that's the source of the error, and returns the range that contains that cell.


## Syntax

 _expression_ . **ShowErrors**

 _expression_ A variable that represents a **Range** object.


### Return Value

Variant


## Example

This example displays a red tracer arrow if there's an error in the active cell on Sheet1.


```vb
Worksheets("Sheet1").Activate 
If IsError(ActiveCell.Value) Then 
 ActiveCell.ShowErrors 
End If
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

