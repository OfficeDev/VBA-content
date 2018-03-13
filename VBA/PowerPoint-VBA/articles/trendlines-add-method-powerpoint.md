---
title: Trendlines.Add Method (PowerPoint)
keywords: vbapp10.chm65717
f1_keywords:
- vbapp10.chm65717
ms.prod: powerpoint
api_name:
- PowerPoint.Trendlines.Add
ms.assetid: d7bd5d75-233f-bdc7-87a4-297b69031838
ms.date: 06/08/2017
---


# Trendlines.Add Method (PowerPoint)

Creates a new trendline.


## Syntax

 _expression_. **Add**( **_Type_**, **_Order_**, **_Period_**, **_Forward_**, **_Backward_**, **_Intercept_**, **_DisplayEquation_**, **_DisplayRSquared_**, **_Name_** )

 _expression_ A variable that represents a **[Trendlines](trendlines-object-powerpoint.md)** object.


### Parameters



| <strong>Name</strong> | <strong>Required/Optional</strong> | <strong>Data Type</strong>                                                                                                                                              | <strong>Description</strong>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                |
|:----------------------|:-----------------------------------|:------------------------------------------------------------------------------------------------------------------------------------------------------------------------|:----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <em>Type</em>         | Optional                           | <strong><a href="xltrendlinetype-enumeration-powerpoint.md" data-raw-source="[XlTrendlineType](xltrendlinetype-enumeration-powerpoint.md)">XlTrendlineType</a></strong> | One of the enumeration values that specifies the trendline type. The default is  <strong>xlLinear</strong>.                                                                                                                                                                                                                                                                                                                                                                                                                                                                 |
| <em>Order</em>        | Optional                           | <strong>Variant</strong>                                                                                                                                                | The trendline order. Required ifType is set to  <strong>xlPolynomial</strong>. If specified, the value must be an integer from 2 through 6.                                                                                                                                                                                                                                                                                                                                                                                                                                 |
| <em>Period</em>       | Optional                           | <strong>Variant</strong>                                                                                                                                                | The trendline period. Required ifType is set to  <strong>xlMovingAvg</strong>. If specified, the value must be an integer greater than 1 and less than the number of data points in the series to which you are adding a trendline.                                                                                                                                                                                                                                                                                                                                         |
| <em>Forward</em>      | Optional                           | <strong>Variant</strong>                                                                                                                                                | The number of periods (or units on a scatter chart) that the trendline extends forward.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                     |
| <em>Backward</em>     | Optional                           | <strong>Variant</strong>                                                                                                                                                | The number of periods (or units on a scatter chart) that the trendline extends backward.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    |
| <em>Intercept</em>    | Optional                           | <strong>Variant</strong>                                                                                                                                                | The trendline intercept. If specified, the value must be a double-precision floating-point number. If omitted, the intercept is automatically set by the regression, and the  <strong><a href="trendline-interceptisauto-property-powerpoint.md" data-raw-source="[InterceptIsAuto](trendline-interceptisauto-property-powerpoint.md)">InterceptIsAuto</a></strong> property of the resulting <strong><a href="trendline-object-powerpoint.md" data-raw-source="[Trendline](trendline-object-powerpoint.md)">Trendline</a></strong> object is set to <strong>True</strong>. |

 **Note**  This parameter is applicable only ifType is set to  **xlExponential**, **xlLinear**, or **xlPolynomial**.

|
| _DisplayEquation_|Optional|**Variant**|**True** to display the equation of the trendline on the chart (in the same data label as the R-squared value). The default is **False**.|
| _DisplayRSquared_|Optional|**Variant**|**True** to display the R-squared value of the trendline on the chart (in the same data label as the equation). The default is **False**.|
| _Name_|Optional|**Variant**|The name of the trendline. If omitted, Microsoft Word generates a name, and the  **[NameIsAuto](trendline-nameisauto-property-powerpoint.md)** property of the resulting **[Trendline](trendline-object-powerpoint.md)** object is set to **True**.|

### Return Value

A  **[Trendline](trendline-object-powerpoint.md)** object that represents the new trendline.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

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


[Trendlines Object](trendlines-object-powerpoint.md)

