---
title: Chart.ApplyDataLabels Method (PowerPoint)
keywords: vbapp10.chm67458
f1_keywords:
- vbapp10.chm67458
ms.prod: powerpoint
api_name:
- PowerPoint.Chart.ApplyDataLabels
ms.assetid: 9d712577-82cc-5d8d-69d1-f5fbaf02c820
ms.date: 06/08/2017
---


# Chart.ApplyDataLabels Method (PowerPoint)

Applies data labels to all the series in a chart.


## Syntax

 _expression_. **ApplyDataLabels**( **_Type_**, **_LegendKey_**, **_AutoText_**, **_HasLeaderLines_**, **_ShowSeriesName_**, **_ShowCategoryName_**, **_ShowValue_**, **_ShowPercentage_**, **_ShowBubbleSize_**, **_Separator_** )

 _expression_ A variable that represents a **[Chart](chart-object-powerpoint.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Optional|**[XlDataLabelsType](xldatalabelstype-enumeration-powerpoint.md)**|One of the enumeration values that specifies the type of data label to apply.|
| _LegendKey_|Optional|**Variant**|**True** to show the legend key next to the point. The default is **False**.|
| _AutoText_|Optional|**Variant**|**True** if the object automatically generates appropriate text based on content.|
| _HasLeaderLines_|Optional|**Variant**|For the  **[Chart](chart-object-powerpoint.md)** and **[Series](series-object-powerpoint.md)** objects, **True** if the series has leader lines.|
| _ShowSeriesName_|Optional|**Variant**|**True** to enable the series name for the data label; otherwise, **False**.|
| _ShowCategoryName_|Optional|**Variant**|**True** to enable the category name for the data label; otherwise, **False**.|
| _ShowValue_|Optional|**Variant**|**True** to enable the value for the data label; otherwise, **False**.|
| _ShowPercentage_|Optional|**Variant**|**True** to enable the percentage for the data label; otherwise, **False**.|
| _ShowBubbleSize_|Optional|**Variant**|**True** to enable the bubble size for the data label; otherwise, **False**.|
| _Separator_|Optional|**Variant**|The separator for the data label.|

## Remarks

The Type parameter can be one of the following  **XlDataLabelsType** constants:


-  **xlDataLabelsShowBubbleSizes** ?The bubble size for the data label.
    
-  **xlDataLabelsShowLabelAndPercent** ?The percentage of the total, and the category for the point. Available only for pie charts and doughnut charts.
    
-  **xlDataLabelsShowPercent** ?The percentage of the total. Available only for pie charts and doughnut charts.
    
-  **xlDataLabelsShowLabel** ?The category for the point.
    
-  **xlDataLabelsShowNone** ?No data labels.
    
-  **xlDataLabelsShowValue** ?(Default) The value for the point (assumed if this argument is not specified).
    

## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example applies category labels to the first series of the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)
    If .HasChart Then
        .Chart.SeriesCollection(1). _
            ApplyDataLabels Type:=xlDataLabelsShowLabel
    End If
End With
```


 **Note**  If you use a string for the Separator parameter, you will get a string as the separator. If you use  **xlDataLabelSeparatorDefault** (= 1), you will get the default data label separator, which is either a comma or a newline, depending on the data label.When a value of 1 is returned, it indicates that the user has not changed the default separator, which is a comma (,). You can also pass a value of 1 to change the separator back to the default separator.The chart must first be active before you can access the data labels programmatically; otherwise, a run-time error occurs.


## See also


#### Concepts


[Chart Object](chart-object-powerpoint.md)

