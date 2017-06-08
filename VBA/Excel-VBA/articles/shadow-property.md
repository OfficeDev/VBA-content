---
title: Shadow Property
keywords: vbagr10.chm65639
f1_keywords:
- vbagr10.chm65639
ms.prod: excel
api_name:
- Excel.Shadow
ms.assetid: 2450bcd9-24fb-57b9-6d73-5ed4abef54d0
ms.date: 06/08/2017
---


# Shadow Property

Shadow property as it applies to the  **AxisTitle**,  **ChartArea**,  **ChartTitle**,  **DataLabel**,  **DataLabels**,  **DisplayUnitLabel**,  **Legend**,  **LegendKey**,  **Point**, and  **Series** objects.

True if the font is a shadow font or if the specified object has a shadow. Read/write Boolean.

 _expression_. **Shadow**

 _expression_ Required. An expression that returns one of the above objects.
Shadow property as it applies to the  **Font** object.
True if the font is a shadow font or if the specified object has a shadow. Read/write Variant.
 _expression_. **Shadow**
 _expression_ Required. An expression that returns one of the above objects.

## Remarks

For the  **Font** object, this property has no effect in Microsoft Windows, but its value is retained (it can be set and returned).


## Example

This example adds a shadow to the title of  `myChart`.


```vb
myChart.ChartTitle.Shadow = True
```


