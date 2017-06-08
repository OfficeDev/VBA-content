---
title: FontStyle Property
keywords: vbagr10.chm65713
f1_keywords:
- vbagr10.chm65713
ms.prod: excel
api_name:
- Excel.FontStyle
ms.assetid: ee63b4bf-1cc1-7348-c79f-c6d4962abe9c
ms.date: 06/08/2017
---


# FontStyle Property

Returns or sets the font style. Read/write Variant.

 _expression_. **FontStyle**

 _expression_ Required. An expression that returns one of the objects in the Applies To list.


## Remarks

Changing this property may affect other  **Font** properties (such as **Bold** and **Italic**).


## Example

This example sets the font style for the chart title to bold and italic.


```
myChart.ChartTitle.Font.FontStyle = "Bold Italic"
```


