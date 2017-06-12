---
title: Chart.SaveChartTemplate Method (PowerPoint)
keywords: vbapp10.chm684008
f1_keywords:
- vbapp10.chm684008
ms.prod: powerpoint
api_name:
- PowerPoint.Chart.SaveChartTemplate
ms.assetid: 568abe18-27d3-4944-7bca-186faa534959
ms.date: 06/08/2017
---


# Chart.SaveChartTemplate Method (PowerPoint)

Saves a custom chart template to the list of available chart templates.


## Syntax

 _expression_. **SaveChartTemplate**( **_FileName_** )

 _expression_ A variable that represents a **[Chart](chart-object-powerpoint.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileName_|Required|**String**|The name of the chart template.|

## Remarks

By default, this method saves the active chart to the user's chart template directory. If a UNC or URL is specified, the chart will be saved to the specified location instead. 


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example adds a new chart template based on the first chart of the active document.




```vb
With ActiveDocument.InlineShapes(1)
    If .HasChart Then
        .Chart.SaveChartTemplate _
            FileName:="Presentation Chart"
    End If
End With
```


## See also


#### Concepts


[Chart Object](chart-object-powerpoint.md)

