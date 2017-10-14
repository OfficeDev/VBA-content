---
title: Series.ErrorBar Method (Excel)
keywords: vbaxl10.chm578081
f1_keywords:
- vbaxl10.chm578081
ms.prod: excel
api_name:
- Excel.Series.ErrorBar
ms.assetid: 0f127c27-09d3-a0e0-7a1d-5e3544039658
ms.date: 06/08/2017
---


# Series.ErrorBar Method (Excel)

Applies error bars to the series.  **Variant** .


## Syntax

 _expression_ . **ErrorBar**( **_Direction_** , **_Include_** , **_Type_** , **_Amount_** , **_MinusValues_** )

 _expression_ A variable that represents a **Series** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Direction_|Required| **[XlErrorBarDirection](xlerrorbardirection-enumeration-excel.md)**|The error bar direction.|
| _Include_|Required| **[XlErrorBarInclude](xlerrorbarinclude-enumeration-excel.md)**|The error bar parts to include.|
| _Type_|Required| **[XlErrorBarType](xlerrorbartype-enumeration-excel.md)**|The error bar type.|
| _Amount_|Optional| **Variant**|The error amount. Used for only the positive error amount when  _Type_ is **xlErrorBarTypeCustom** .|
| _MinusValues_|Optional| **Variant**|The negative error amount when  _Type_ is **xlErrorBarTypeCustom** .|

### Return Value

Variant


## Example

This example applies standard error bars in the Y direction for series one in Chart1. The error bars are applied in the positive and negative directions. The example should be run on a 2-D line chart.


```vb
Charts("Chart1").SeriesCollection(1).ErrorBar _ 
 Direction:=xlY, Include:=xlErrorBarIncludeBoth, _ 
 Type:=xlErrorBarTypeStError
```


## See also


#### Concepts


[Series Object](series-object-excel.md)

