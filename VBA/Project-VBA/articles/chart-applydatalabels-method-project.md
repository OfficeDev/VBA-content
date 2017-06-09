---
title: Chart.ApplyDataLabels Method (Project)
ms.prod: project-server
ms.assetid: cda031a4-ed86-1ec8-583d-44767785e3a1
ms.date: 06/08/2017
---


# Chart.ApplyDataLabels Method (Project)
Applies data labels to all the series in a chart.

## Syntax

 _expression_. **ApplyDataLabels** _(Type,_ _IMsoLegendKey,_ _AutoText,_ _HasLeaderLines,_ _ShowSeriesName,_ _ShowCategoryName,_ _ShowValue,_ _ShowPercentage,_ _ShowBubbleSize,_ _Separator)_

 _expression_ A variable that represents a **Chart** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Optional|**Office.XlDataLabelsType**|The type of data label to apply. The default value is  **xlDataLabelsShowValue**.|
| _IMsoLegendKey_|Optional|**Variant**|**True** to show the legend key next to the point. The default value is **False**.|
| _AutoText_|Optional|**Variant**|**True** if the object automatically generates appropriate text based on content.|
| _HasLeaderLines_|Optional|**Variant**|**True** if the series has leader lines.|
| _ShowSeriesName_|Optional|**Variant**|**True** to enable the series name for the data label. **False** to disable the series name.|
| _ShowCategoryName_|Optional|**Variant**|**True** to enable the category name for the data label. **False** to disable the category name.|
| _ShowValue_|Optional|**Variant**|**True** to enable the value for the data label. **False** to disable the value.If the  _Type_ parameter is not specified, _ShowValue_ is assumed to be **True**.|
| _ShowPercentage_|Optional|**Variant**|**True** to enable the percentage for the data label. **False** to disable the percentage.|
| _ShowBubbleSize_|Optional|**Variant**|**True** to enable the bubble size for the data label. **False** to disable the bubble size.|
| _Separator_|Optional|**Variant**|The separator for the data label.|

### Return value

 **Nothing**


## Example

The following example applies data labels to each data point.


```vb
Sub SetDataLabels()
    Dim chartShape As Shape
    Dim reportName As String
    
    reportName = "Simple scalar chart"
    Set chartShape = ActiveProject.Reports(reportName).Shapes(1)
    
    chartShape.Chart.ApplyDataLabels
End Sub
```


## See also


#### Other resources


[Chart Object](chart-object-project.md)
