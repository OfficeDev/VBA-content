---
title: Application.InchesToPoints Method (Excel)
keywords: vbaxl10.chm133148
f1_keywords:
- vbaxl10.chm133148
ms.prod: excel
api_name:
- Excel.Application.InchesToPoints
ms.assetid: 7689eae4-f533-32e3-d431-4873029a8bc1
ms.date: 06/08/2017
---


# Application.InchesToPoints Method (Excel)

Converts a measurement from inches to points.


## Syntax

 _expression_ . **InchesToPoints**( **_Inches_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Inches_|Required| **Double**|Specifies the inch value to be converted to points.|

### Return Value

Double


## Example

This example sets the left margin of Sheet1 to 2.5 inches.


```vb
Worksheets("Sheet1").PageSetup.LeftMargin = _ 
 Application.InchesToPoints(2.5)
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

