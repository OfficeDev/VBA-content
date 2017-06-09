---
title: Trendlines.Add Method (Excel)
keywords: vbaxl10.chm592074
f1_keywords:
- vbaxl10.chm592074
ms.prod: excel
api_name:
- Excel.Trendlines.Add
ms.assetid: 4d86029e-3c42-2d81-69d3-94d8dc072ccd
ms.date: 06/08/2017
---


# Trendlines.Add Method (Excel)

Creates a new trendline.


## Syntax

 _expression_ . **Add**( **_Type_** , **_Order_** , **_Period_** , **_Forward_** , **_Backward_** , **_Intercept_** , **_DisplayEquation_** , **_DisplayRSquared_** , **_Name_** )

 _expression_ A variable that represents a **Trendlines** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Optional| **[XlTrendlineType](xltrendlinetype-enumeration-excel.md)**|The trendline type.|
| _Order_|Optional| **Variant**| **Variant** . if _Type_ is **xlPolynomial** . The trendline order. Must be an integer from 2 to 6, inclusive.|
| _Period_|Optional| **Variant**|if  _Type_ is **xlMovingAvg** . The trendline period. Must be an integer greater than 1 and less than the number of data points in the series you're adding a trendline to.|
| _Forward_|Optional| **Variant**|The number of periods (or units on a scatter chart) that the trendline extends forward.|
| _Backward_|Optional| **Variant**|The number of periods (or units on a scatter chart) that the trendline extends backward.|
| _Intercept_|Optional| **Variant**|The trendline intercept. If this argument is omitted, the intercept is automatically set by the regression.|
| _DisplayEquation_|Optional| **Variant**| **True** to display the equation of the trendline on the chart (in the same data label as the R-squared value). The default value is **False** .|
| _DisplayRSquared_|Optional| **Variant**| **True** to display the R-squared value of the trendline on the chart (in the same data label as the equation). The default value is **False** .|
| _Name_|Optional| **Variant**|The name of the trendline as text. If this argument is omitted, Microsoft Excel generates a name.|

### Return Value

A  **[Trendline](trendline-object-excel.md)** object that represents the new trendline.


## Example

This example creates a new linear trendline in Chart1.


```vb
ActiveWorkbook.Charts("Chart1").SeriesCollection(1).Trendlines.Add
```


## See also


#### Concepts


[Trendlines Object](trendlines-object-excel.md)

