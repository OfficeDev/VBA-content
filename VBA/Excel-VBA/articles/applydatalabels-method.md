---
title: ApplyDataLabels Method
keywords: vbagr10.chm67458
f1_keywords:
- vbagr10.chm67458
ms.prod: excel
api_name:
- Excel.ApplyDataLabels
ms.assetid: 1750d716-66f8-fe4e-8023-fbcfcc5c5ff5
ms.date: 06/08/2017
---


# ApplyDataLabels Method

ApplyDataLabels method as it applies to the  **Chart** object.

Applies data labels to a point, a series, or all the series in a chart.

 _expression_. **ApplyDataLabels**( **_Type_**,  **_LegendKey_**,  **_AutoText_**,  **_HasLeaderLines_**)

 _expression_ Required. An expression that returns a **Chart** object.
 **Type**Optional 
 **Variant**
. The data label type.


|XlDataLabelsType can be one of these XlDataLabelsType constants.|
| **xlDataLabelsShowBubbleSizes** The bubble size for the data label.|
| **xlDataLabelsShowLabel**. Category for the point.|
| **xlDataLabelsShowLabelAndPercent**. Percentage of the total, and category for the point. Available only for pie charts and doughnut charts.|
| **xlDataLabelsShowNone**. No data labels.|
| **xlDataLabelsShowPercent**. Percentage of the total. Available only for pie charts and doughnut charts.|
| **xlDataLabelsShowValue**_default_. Value for the point.|
 **LegendKey** Optional **Variant**.  **True** to show the legend key next to the point. The default value is **False**.
 **AutoText** Optional **Variant**.  **True** if the object automatically generates appropriate text based on content.
 **HasLeaderLines**Optional  **Variant**.  **True** if the series has leader lines.
ApplyDataLabels method as it applies to the  **Point** and **Series** objects.
Applies data labels to a point, a series, or all the series in a chart.
 _expression_. **ApplyDataLabels**( **_Type_**,  **_LegendKey_**,  **_AutoText_**,  **_HasLeaderLines_**,  **_ShowSeriesName_**,  **_ShowCategoryName_**,  **_ShowValue_**,  **_ShowPercentage_**,  **_ShowBubbleSize_**,  **_Separator_**)
 _expression_ Required. An expression that returns one of the above objects.
 **Type**Optional 
 **XlDataLabelsType**
. The data label type.


|XlDataLabelsType can be one of these XlDataLabelsType constants.|
| **xlDataLabelsShowBubbleSizes** The bubble size for the data label.|
| **xlDataLabelsShowLabel**. Category for the point.|
| **xlDataLabelsShowLabelAndPercent**. Percentage of the total, and category for the point. Available only for pie charts and doughnut charts.|
| **xlDataLabelsShowNone**. No data labels.|
| **xlDataLabelsShowPercent**. Percentage of the total. Available only for pie charts and doughnut charts.|
| **xlDataLabelsShowValue**_default_. Value for the point.|
 **LegendKey** Optional **Variant**.  **True** to show the legend key next to the point. The default value is **False**.
 **AutoText** Optional **Variant**.  **True** if the object automatically generates appropriate text based on content.
 **HasLeaderLines**Optional  **Variant**.  **True** if the series has leader lines.
 **ShowSeriesName**Optional  **Variant**. The series name for the data label.
 **ShowCategoryName**Optional  **Variant**. The category name for the data label.
 **ShowValue**Optional  **Variant**. The value for the data label.
 **ShowPercentage**Optional  **Variant**. The percentage for the data label.
 **ShowBubbleSize**Optional  **Variant**. The bubble size for the data label.
 **Separator**Optional  **Variant**. The separator for the data label.

## Example

ApplyDataLabels method as it applies to the  **Series** object.

This example applies category labels to series one.




```
myChart.SeriesCollection(1). _ 
 ApplyDataLabels Type:=xlDataLabelsShowLabel
```


