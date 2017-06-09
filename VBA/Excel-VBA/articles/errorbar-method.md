---
title: ErrorBar Method
keywords: vbagr10.chm65688
f1_keywords:
- vbagr10.chm65688
ms.prod: excel
api_name:
- Excel.ErrorBar
ms.assetid: c2ada146-1549-aa88-2a39-bf1cccf1008b
ms.date: 06/08/2017
---


# ErrorBar Method

Applies error bars to the specified series. Variant.

 _expression_. **ErrorBar**( **_Direction_**,  **_Include_**,  **_Type_**,  **_Amount_**,  **_MinusValues_**)

 _expression_ Required. An expression that returns one of the objects in the Applies To list.

 **Direction**Required 
 **XlErrorBarDirection**
. The error bar direction.


|XlErrorBarDirection can be one of these XlErrorBarDirection constants.|
| **xlX** Can only be used with scatter charts.|
| **xlY**_default._|
 **Include**Required 
 **XlErrorBarInclude**
. The error bar parts to be included.


|XlErrorBarInclude can be one of these XlErrorBarInclude constants.|
| **xlErrorBarIncludeBoth**_default._|
| **xlErrorBarIncludeMinusValues**|
| **xlErrorBarIncludeNone**|
| **xlErrorBarIncludePlusValues**|
 **Type**Required 
 **XlErrorBarType**
. The error bar type.


|XlErrorBarType can be one of these XlErrorBarType constants.|
| **xlErrorBarTypeCustom**|
| **xlErrorBarTypeFixedValue**|
| **xlErrorBarTypePercent**|
| **xlErrorBarTypeStDev**|
| **xlErrorBarTypeStError**|
 **Amount** Optional **Variant**. The error amount. Used for only the positive error amount when  **_Type_** is **xlErrorBarTypeCustom**.
 **MinusValues** Optional **Variant**. The negative error amount when  **_Type_** is **xlErrorBarTypeCustom**.

## Example

This example applies standard error bars in the Y direction for series one. The error bars are applied in the positive and negative directions. The example should be run on a 2-D line chart.


```
myChart.SeriesCollection(1).ErrorBar _ 
 Direction:=xlY, Include:=xlErrorBarIncludeBoth, _ 
 Type:=xlErrorBarTypeStError
```


