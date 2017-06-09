---
title: Chart.SeriesCollection Method (Project)
keywords: vbapj.chm131631
f1_keywords:
- vbapj.chm131631
ms.prod: project-server
ms.assetid: fb4fea11-3dac-73f9-6566-6c81de0888e7
ms.date: 06/08/2017
---


# Chart.SeriesCollection Method (Project)
Returns an object that represents either one series (a  **[Series](series-object-project.md)** object) or a collection of the series (a **[SeriesCollection](seriescollection-object-project.md)** object) in the chart or chart group.

## Syntax

 _expression_. **SeriesCollection** _(Index)_

 _expression_ A variable that represents a **Chart** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Optional|**Variant**|The name or index number of the series. If  _Index_ is not specified, the **SeriesCollection** method returns all of the series in the chart.|
| _Index_|Optional|VARIANT||

### Return value

 **Object**


## Example

To get a single series, specify the  _Index_ parameter. The following example prints the first value of the "Actual Work" series. The first call to the **SeriesCollection** method gets the collection of all the series in the chart. The second call to the **SeriesCollection** method gets one specific series.


```vb
Sub GetSeriesValue()
    Dim reportName As String
    Dim theReportIndex As Integer
    Dim theChart As Chart
    Dim seriesInChart As SeriesCollection
    Dim chartSeries As Series
    
    reportName = "Simple scalar chart"
        
    If (ActiveProject.Reports.IsPresent(reportName)) Then
        ' Make the report active.
        theReportIndex = ActiveProject.Reports(reportName).Index
        ActiveProject.Reports(theReportIndex).Apply
        
        Set theChart = ActiveProject.Reports(theReportIndex).Shapes(1).Chart
        Set seriesInChart = theChart.SeriesCollection
        
        If (seriesInChart.Count > 1) Then
            Set chartSeries = theChart.SeriesCollection("Actual Work")
            Debug.Print "Value of the Actual Work series, for task " &; chartSeries.XValues(1) _
                &; ": " &; chartSeries.Values(1)
        End If
        
    End If
End Sub
```

For example, running the  **GetSeriesValue** macro on a chart that includes a plot of actual work for tasks, could have the following output: `Value of the Actual Work series, for task T1: 16`


## See also


#### Other resources


[Chart Object](chart-object-project.md)
