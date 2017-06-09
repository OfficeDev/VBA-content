---
title: Series.ApplyDataLabels Method (Excel)
keywords: vbaxl10.chm578122
f1_keywords:
- vbaxl10.chm578122
ms.prod: excel
api_name:
- Excel.Series.ApplyDataLabels
ms.assetid: 959a4d12-ed48-48fc-04cf-7a1880cd7e1f
ms.date: 06/08/2017
---


# Series.ApplyDataLabels Method (Excel)

Applies data labels to a series.


## Syntax

 _expression_ . **ApplyDataLabels**( **_Type_** , **_LegendKey_** , **_AutoText_** , **_HasLeaderLines_** , **_ShowSeriesName_** , **_ShowCategoryName_** , **_ShowValue_** , **_ShowPercentage_** , **_ShowBubbleSize_** , **_Separator_** )

 _expression_ A variable that represents a **Series** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Optional| **[XlDataLabelsType](xldatalabelstype-enumeration-excel.md)**|The type of data label to apply.|
| _LegendKey_|Optional| **Variant**| **True** to show the legend key next to the point. The default value is **False** .|
| _AutoText_|Optional| **Variant**| **True** if the object automatically generates appropriate text based on content.|
| _HasLeaderLines_|Optional| **Variant**|For the  **[Chart](chart-object-excel.md)** and **[Series](series-object-excel.md)** objects, **True** if the series has leader lines.|
| _ShowSeriesName_|Optional| **Variant**|Pass a boolean value to enable or disable the series name for the data label.|
| _ShowCategoryName_|Optional| **Variant**|Pass a boolean value to enable or disable the category name for the data label.|
| _ShowValue_|Optional| **Variant**|Pass a boolean value to enable or disable the value for the data label.|
| _ShowPercentage_|Optional| **Variant**|Pass a boolean value to enable or disable the percentage for the data label.|
| _ShowBubbleSize_|Optional| **Variant**|Pass a boolean value to enable or disable the bubble size for the data label.|
| _Separator_|Optional| **Variant**|The separator for the data label.|

## Remarks





| **XlDataLabelsType** can be one of these **XlDataLabelsType** constants.|
| **xlDataLabelsShowBubbleSizes** . The bubble size for the data label.|
| **xlDataLabelsShowLabelAndPercent** . Percentage of the total, and category for the point. Available only for pie charts and doughnut charts.|
| **xlDataLabelsShowPercent** . Percentage of the total. Available only for pie charts and doughnut charts.|
| **xlDataLabelsShowLabel** . Category for the point.|
| **xlDataLabelsShowNone** . No data labels.|
| **xlDataLabelsShowValue** . _default_ . Value for the point (assumed if this argument isn't specified).|

## Example

This example applies category labels to series one in Chart1.


```vb
Charts("Chart1").SeriesCollection(1). _ 
 ApplyDataLabels Type:=xlDataLabelsShowLabel
```


## See also


#### Concepts


[Series Object](series-object-excel.md)

