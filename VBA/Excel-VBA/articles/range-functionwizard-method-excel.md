---
title: Range.FunctionWizard Method (Excel)
keywords: vbaxl10.chm144139
f1_keywords:
- vbaxl10.chm144139
ms.prod: excel
api_name:
- Excel.Range.FunctionWizard
ms.assetid: a9a0c765-4903-4969-8f09-c8f051213a96
ms.date: 06/08/2017
---


# Range.FunctionWizard Method (Excel)

Starts the Function Wizard for the upper-left cell of the range.


## Syntax

 _expression_ . **FunctionWizard**

 _expression_ A variable that represents a **Range** object.


### Return Value

Variant


## Example

This example starts the Function Wizard for the active cell on Sheet1.


```vb
Worksheets("Sheet1").Activate 
ActiveCell.FunctionWizard
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

