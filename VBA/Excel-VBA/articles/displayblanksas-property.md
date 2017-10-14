---
title: DisplayBlanksAs Property
keywords: vbagr10.chm3077021
f1_keywords:
- vbagr10.chm3077021
ms.prod: excel
api_name:
- Excel.DisplayBlanksAs
ms.assetid: c2669ad5-9532-ea7c-120c-bc8a15878864
ms.date: 06/08/2017
---


# DisplayBlanksAs Property

Returns or sets the way that blank cells are plotted on a chart. Read/write XlDisplayBlanksAs .



|XlDisplayBlanksAs can be one of these XlDisplayBlanksAs constants.|
| **xlInterpolated**|
| **xlNotPlotted**|
| **xlZero**|

 _expression_. **DisplayBlanksAs**

 _expression_ Required. An expression that returns one of the objects in the Applies To list.

## Example

This example sets Microsoft Graph to not plot blank cells.


```
myChart.DisplayBlanksAs = xlNotPlotted
```


