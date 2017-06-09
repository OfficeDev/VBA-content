---
title: Add Method
keywords: vbagr10.chm3077604
f1_keywords:
- vbagr10.chm3077604
ms.prod: excel
api_name:
- Excel.Add
ms.assetid: 529bbd0e-c726-2e88-fa75-d492fede7f37
ms.date: 06/08/2017
---


# Add Method

Creates a new trend line. Returns a Trendline object.

 _expression_. **Add**( **_Type_**,  **_Order_**,  **_Period_**,  **_Forward_**,  **_Backward_**,  **_Intercept_**,  **_DisplayEquation_**,  **_DisplayRSquared_**,  **_Name_**)

 _expression_ Required. An expression that returns one of the objects in the Applies To list.

 **Type**Optional 
 **XlTrendlineType**
. The type of trendline.


|XlTrendlineType can be one of these XlTrendlineType constants.|
| **xlExponential**|
| **xlLinear**_default_|
| **xlLogarithmic**|
| **xlMovingAvg**|
| **xlPolynomial**|
| **xlPower**|
 **Order** Optional **Variant**. Required if  **_Type_** is **xlPolynomial**. The trendline order. Must be an integer from 2 through 6.
 **Period** Optional **Variant**. Required if  **_Type_** is **xlMovingAvg**. The trendline period. Must be an integer greater than 1 and less than the number of data points in the series you're adding a trendline to.
 **Forward** Optional **Variant**. The number of periods (or units on a scatter chart) that the trendline extends forward.
 **Backward** Optional **Variant**. The number of periods (or units on a scatter chart) that the trendline extends backward.
 **Intercept** Optional **Variant**. The trendline intercept. If this argument is omitted, the intercept is automatically set by the regression.
 **DisplayEquation** Optional **Variant**.  **True** to display the equation of the trendline on the chart (in the same data label as the R-squared value). The default value is **False**.
 **DisplayRSquared** Optional **Variant**.  **True** to display the R-squared value of the trendline on the chart (in the same data label as the equation). The default value is **False**.
 **Name** Optional **Variant**. The name of the trendline, as text. If this argument is omitted, Microsoft Graph generates a name.

## Example

This example creates a new linear trendline on the chart.


```
myChart.SeriesCollection(1).Trendlines.Add
```


