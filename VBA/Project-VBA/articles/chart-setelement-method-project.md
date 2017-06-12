---
title: Chart.SetElement Method (Project)
ms.prod: project-server
ms.assetid: ca4acf62-c090-f11c-2816-c5e1a75762fa
ms.date: 06/08/2017
---


# Chart.SetElement Method (Project)
Adds the specified element to a chart or to a selected object on a chart.

## Syntax

 _expression_. **SetElement** _(RHS)_

 _expression_ A variable that represents a **Chart** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
| _RHS_|Required|**MsoChartElementType**|One of the enumeration constants for the chart element type to add.|

### Return value

 **Nothing**


## Remarks

The  _RHS_ value for the **SetElement** method corresponds to items in the **Add Chart Element** submenus. Different items are enabled, depending on the type of chart. If you try to add an element that does not exist for a particular chart, you get an unspecified error. For example, on a 3-D chart, the **Error Bars** item in the **Add Chart Element** drop-down list is unavailable. A call to `Chart.SetElement msoElementErrorBarStandardDeviation` results in an error.


## Example

The following example adds minor gridlines to the value axis, and adds data label callouts to the second data series.


```vb
Sub TestSetElements()
    Dim chartShape As Shape
    Dim reportName As String
    
    reportName = "Simple 3-D chart"
    Set chartShape = ActiveProject.Reports(reportName).Shapes(1)
    
    With chartShape.Chart
        .SetElement msoElementChartTitleAboveChart
        
        ' Select the major gridlines on the value axis, and then add minor gridlines.
        .axes(Office.xlValue).MajorGridlines.Select
        .SetElement msoElementPrimaryCategoryGridLinesMinor
        
        ' Select the second data series and add data label callouts.
        If .SeriesCollection.Count > 1 Then
            .SeriesCollection(2).Select
            .SetElement msoElementDataLabelCallout
        End If
    End With
End Sub
```


## See also


#### Other resources


[Chart Object](chart-object-project.md)
