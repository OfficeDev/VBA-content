---
title: SetDefaultChart Method
keywords: vbagr10.chm65755
f1_keywords:
- vbagr10.chm65755
ms.prod: excel
api_name:
- Excel.SetDefaultChart
ms.assetid: 1afc1023-654b-67cd-aace-bc4b87747520
ms.date: 06/08/2017
---


# SetDefaultChart Method

Specifies the name of the chart template that Microsoft Graph will use when creating new charts.

 _expression_. **SetDefaultChart**( **_FormatName_**,  **_Gallery_**)

 _expression_ Required. An expression that returns one of the objects in the Applies To list.

 _FormatName_ Optional **Variant**. The name of the specified custom autoformat. This name can be a string that denotes the custom autoformat, or it can be the special constant  **xlBuiltIn** to specify the built-in chart template.
 **Gallery**Optional  **Variant**.

## Example

This example sets the default chart template to the custom autoformat named "Monthly Sales."


```
myChart.Application.SetDefaultChart FormatName:="Monthly Sales"
```


