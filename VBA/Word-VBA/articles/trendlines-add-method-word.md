---
title: Trendlines.Add Method (Word)
keywords: vbawd10.chm102367413
f1_keywords:
- vbawd10.chm102367413
ms.prod: word
api_name:
- Word.Trendlines.Add
ms.assetid: 7260373c-626b-2778-0517-e5c62b754bc9
ms.date: 06/08/2017
---


# Trendlines.Add Method (Word)

Creates a new trendline.


## Syntax

 _expression_ . **Add**( **_Type_** , **_Order_** , **_Period_** , **_Forward_** , **_Backward_** , **_Intercept_** , **_DisplayEquation_** , **_DisplayRSquared_** , **_Name_** )

 _expression_ A variable that represents a **[Trendlines](trendlines-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Optional| **[XlTrendlineType](xltrendlinetype-enumeration-word.md)**|One of the enumeration values that specifies the trendline type. The default is  **xlLinear** .|
| _Order_|Optional| **Variant**|The trendline order. Required ifType is set to  **xlPolynomial** . If specified, the value must be an integer from 2 through 6.|
| _Period_|Optional| **Variant**|The trendline period. Required ifType is set to  **xlMovingAvg** . If specified, the value must be an integer greater than 1 and less than the number of data points in the series to which you are adding a trendline.|
| _Forward_|Optional| **Variant**|The number of periods (or units on a scatter chart) that the trendline extends forward.|
| _Backward_|Optional| **Variant**|The number of periods (or units on a scatter chart) that the trendline extends backward.|
| _Intercept_|Optional| **Variant**|The trendline intercept. If specified, the value must be a double-precision floating-point number. If omitted, the intercept is automatically set by the regression, and the  **[InterceptIsAuto](trendline-interceptisauto-property-word.md)** property of the resulting **[Trendline](trendline-object-word.md)** object is set to **True** .
 **Note**  This parameter is applicable only ifType is set to  **xlExponential** , **xlLinear** , or **xlPolynomial** .

|
| _DisplayEquation_|Optional| **Variant**| **True** to display the equation of the trendline on the chart (in the same data label as the R-squared value). The default is **False** .|
| _DisplayRSquared_|Optional| **Variant**| **True** to display the R-squared value of the trendline on the chart (in the same data label as the equation). The default is **False** .|
| _Name_|Optional| **Variant**|The name of the trendline. If omitted, Microsoft Word generates a name, and the  **[NameIsAuto](trendline-nameisauto-property-word.md)** property of the resulting **[Trendline](trendline-object-word.md)** object is set to **True** .|

### Return Value

A  **[Trendline](trendline-object-word.md)** object that represents the new trendline.


## Example

The following example creates a new linear trendline for the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
    If .HasChart Then 
        .Chart.SeriesCollection(1).Trendlines.Add 
    End If 
End With
```


## See also


#### Concepts


[Trendlines Object](trendlines-object-word.md)

