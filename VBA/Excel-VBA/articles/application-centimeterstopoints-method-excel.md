---
title: Application.CentimetersToPoints Method (Excel)
keywords: vbaxl10.chm133090
f1_keywords:
- vbaxl10.chm133090
ms.prod: excel
api_name:
- Excel.Application.CentimetersToPoints
ms.assetid: 2693973c-7d80-8883-6959-afabdb51b9b2
ms.date: 06/08/2017
---


# Application.CentimetersToPoints Method (Excel)

Converts a measurement from centimeters to points (one point equals 0.035 centimeters).


## Syntax

 _expression_ . **CentimetersToPoints**( **_Centimeters_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Centimeters_|Required| **Double**|Specifies the centimeter value to be converted to points.|

### Return Value

Double


## Example

This example sets the left margin of Sheet1 to 5 centimeters.


```vb
Worksheets("Sheet1").PageSetup.LeftMargin = _ 
 Application.CentimetersToPoints(5)
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

