---
title: Series.ApplyDataLabels Method (Word)
keywords: vbawd10.chm123733890
f1_keywords:
- vbawd10.chm123733890
ms.prod: word
api_name:
- Word.Series.ApplyDataLabels
ms.assetid: f172d97a-53a9-929d-b929-bfea03d38b91
ms.date: 06/08/2017
---


# Series.ApplyDataLabels Method (Word)

Applies data labels to a series.


## Syntax

 _expression_ . **ApplyDataLabels**( **_Type_** , **_LegendKey_** , **_AutoText_** , **_HasLeaderLines_** , **_ShowSeriesName_** , **_ShowCategoryName_** , **_ShowValue_** , **_ShowPercentage_** , **_ShowBubbleSize_** , **_Separator_** )

 _expression_ A variable that represents a **[Series](series-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Optional| **[XlDataLabelsType](xldatalabelstype-enumeration-word.md)**|The type of data label to apply.|
| _LegendKey_|Optional| **Variant**| **True** to show the legend key next to the point. The default is **False** .|
| _AutoText_|Optional| **Variant**| **True** if the object automatically generates appropriate text based on content.|
| _HasLeaderLines_|Optional| **Variant**|For the  **[Chart](chart-object-word.md)** and **[Series](series-object-word.md)** objects, **True** if the series has leader lines.|
| _ShowSeriesName_|Optional| **Variant**| **True** to enable the series name for the data label; otherwise, **False** .|
| _ShowCategoryName_|Optional| **Variant**| **True** to enable the category name for the data label; otherwise, **False** .|
| _ShowValue_|Optional| **Variant**| **True** to enable the value for the data label; otherwise, **False** .|
| _ShowPercentage_|Optional| **Variant**| **True** to enable the percentage for the data label; otherwise, **False** .|
| _ShowBubbleSize_|Optional| **Variant**| **True** to enable the bubble size for the data label; otherwise, **False** .|
| _Separator_|Optional| **Variant**|The separator for the data label.|

## Remarks

The Type parameter can be one of the following  **XlDataLabelsType** constants:


-  **xlDataLabelsShowBubbleSizes** ?The bubble size for the data label.
    
-  **xlDataLabelsShowLabelAndPercent** ?The percentage of the total, and the category for the point. Available only for pie charts and doughnut charts.
    
-  **xlDataLabelsShowPercent** ?The percentage of the total. Available only for pie charts and doughnut charts.
    
-  **xlDataLabelsShowLabel** ?The category for the point.
    
-  **xlDataLabelsShowNone** ?No data labels.
    
-  **xlDataLabelsShowValue** ?(Default) The value for the point (assumed if this argument is not specified).
    
- 
    

## Example

The following example applies category labels to series one of the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1). _ 
 ApplyDataLabels Type:=xlDataLabelsShowLabel 
 End If 
End With
```


## See also


#### Concepts


[Series Object](series-object-word.md)

