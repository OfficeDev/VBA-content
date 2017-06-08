---
title: Underline Property
keywords: vbagr10.chm65642
f1_keywords:
- vbagr10.chm65642
ms.prod: excel
api_name:
- Excel.Underline
ms.assetid: 82eb4816-bf37-8a6c-046c-a38ea5c275c2
ms.date: 06/08/2017
---


# Underline Property

Returns or sets the type of underline applied to the font. Required  **XlUnderlineStyle**.



|XlUnderlineStyle can be one of these XlUnderlineStyle constants.|
| **xlUnderlineStyleNone**|
| **xlUnderlineStyleSingle** **xlUnderlineStyleDouble** **xlUnderlineStyleSingleAccounting** **xlUnderlineStyleDoubleAccounting**|

 _expression_. **Underline**

 _expression_ Required. An expression that returns one of the objects in the Applies To list.

## Example

This example sets the font in the chart title to single underline.


```
myChart.ChartTitle.Font.Underline = xlUnderlineStyleSingle
```


