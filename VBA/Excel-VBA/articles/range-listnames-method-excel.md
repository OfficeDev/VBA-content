---
title: Range.ListNames Method (Excel)
keywords: vbaxl10.chm144155
f1_keywords:
- vbaxl10.chm144155
ms.prod: excel
api_name:
- Excel.Range.ListNames
ms.assetid: 0523f9b3-d422-76b6-889c-75619cb5b9a6
ms.date: 06/08/2017
---


# Range.ListNames Method (Excel)

Pastes a list of all nonhidden names onto the worksheet, beginning with the first cell in the range.


## Syntax

 _expression_ . **ListNames**

 _expression_ A variable that represents a **Range** object.


### Return Value

Variant


## Remarks

Use the  **[Names](worksheet-names-property-excel.md)** property to return a collection of all the names on a worksheet.


## Example

This example pastes a list of defined names into cell A1 on Sheet1. The example pastes both workbook-level names and sheet-level names defined on Sheet1.


```vb
Worksheets("Sheet1").Range("A1").ListNames
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

