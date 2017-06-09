---
title: Chart.GetChartElement Method (Project)
keywords: vbapj.chm131624
f1_keywords:
- vbapj.chm131624
ms.prod: project-server
ms.assetid: f2705f1d-7252-41ec-848b-f7f9cc26663e
ms.date: 06/08/2017
---


# Chart.GetChartElement Method (Project)
Returns information about the chart element at specified X and Y coordinates. This method will be removed in the released version of Project 2013.

## Syntax

 _expression_. **GetChartElement** _(x,_? _y,_? _ElementID,_? _Arg1,_? _Arg2)_

 _expression_ A variable that represents a **Chart** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
| _x_|Required|**Long**|The X coordinate of the chart element.|
| _y_|Required|**Long**|The Y coordinate of the chart element.|
| _ElementID_|Required|**Long**|When the  **GetChartElement** method returns, _ElementID_ contains the **Office.XLChartItem** value of the chart element at the specified coordinates. For more information, see[Remarks](#pj15_VBAGetChartElement_Remarks).|
| _Arg1_|Required|**Long**|When the method returns,  _Arg1_ contains information related to the chart element. For more information, see[Remarks](#pj15_VBAGetChartElement_Remarks).|
| _Arg2_|Required|**Long**|When the method returns,  _Arg2_ contains information related to the chart element. For more information, see[Remarks](#pj15_VBAGetChartElement_Remarks).|
| _x_|Required|INT32||
| _y_|Required|INT32||
| _ElementID_|Required|INT32||
| _Arg1_|Required|INT32||
| _Arg2_|Required|INT32||

### Return value

The  **GetChartElement** method returns **Nothing**. Returned values are in the  _ElementID_,  _Arg1_, and  _Arg2_ parameters.


## Remarks
<a name="pj15_VBAGetChartElement_Remarks"> </a>


 **Note**  The  **GetChartElement** method will be removed in the released version of Project 2013. The **Chart** object in Project does not implement events; so, a chart in Project cannot be animated with the **GetChartElement** method by interacting with mouse events.

The  **GetChartElement** method is unusual because you specify values for only the first two arguments. Project returns data in the other arguments, and your code should examine those values when the method returns.

The value of  _ElementID_ after the method returns determines whether _Arg1_ and _Arg2_ contain any information (see Table 1).


**Table 1. Information in Arg1 and Arg2, based on the element ID**


|**ElementID Constant**|**Constant Value **|**Arg1**|**Arg2**|
|:-----|:-----|:-----|:-----|
|**xlAxis**|21|AxisIndex|AxisType|
|**xlAxisTitle**|17|AxisIndex|AxisType|
|**xlDisplayUnitLabel**|30|AxisIndex|AxisType|
|**xlMajorGridlines**|15|AxisIndex|AxisType|
|**xlMinorGridlines**|16|AxisIndex|AxisType|
|**xlPivotChartDropZone**|32|DropZoneType|None|
|**xlPivotChartFieldButton**|31|DropZoneType|PivotFieldIndex|
|**xlDownBars**|20|GroupIndex|None|
|**xlDropLines**|26|GroupIndex|None|
|**xlHiLoLines**|25|GroupIndex|None|
|**xlRadarAxisLabels**|27|GroupIndex|None|
|**xlSeriesLines**|22|GroupIndex|None|
|**xlUpBars**|18|GroupIndex|None|
|**xlChartArea**|2|None|None|
|**xlChartTitle**|4|None|None|
|**xlCorners**|6|None|None|
|**xlDataTable**|7|None|None|
|**xlFloor**|23|None|None|
|**xlLeaderLines**|29|None|None|
|**xlLegend**|24|None|None|
|**xlNothing**|28|None|None|
|**xlPlotArea**|19|None|None|
|**xlWalls**|5|None|None|
|**xlDataLabel**|7|SeriesIndex|PointIndex|
|**xlErrorBars**|9|SeriesIndex|None|
|**xlLegendEntry**|12|SeriesIndex|None|
|**xlLegendKey**|13|SeriesIndex|None|
|**xlSeries**|3|SeriesIndex|PointIndex|
|**xlShape**|14|ShapeIndex|None|
|**xlTrendline**|8|SeriesIndex|TrendLineIndex|
|**xlXErrorBars**|10|SeriesIndex|None|
|**xlYErrorBars**|11|SeriesIndex|None|
?

Table 2 describes the meaning of  _Arg1_ and _Arg2_ after the method returns. Values in the **Argument** column are from Table 1.


**Table 2. Meaning of data in Arg1 and Arg2**


|**Argument**|**Description**|
|:-----|:-----|
|AxisIndex|Specifies whether the axis is primary or secondary. Can be one of the following  **Office.XlAxisGroup** constants: **xlPrimary** or **xlSecondary**.|
|AxisType|Specifies the axis type. Can be one of the following  **Office.XlAxisType** constants: **xlCategory**,  **xlSeriesAxis**, or  **xlValue**.|
|DropZoneType|Specifies the drop zone type: column, data, page, or row field. Can be one of the following  **Office.XlPivotFieldOrientation** constants: **xlColumnField**,  **xlDataField**,  **xlPageField**, or  **xlRowField**. The column and row field constants specify the series and category fields, respectively.|
|GroupIndex|Specifies the offset within the  **Office.IMsoChart.ChartGroups** collection for a specific chart group.|
|PivotFieldIndex|Specifies the offset within the  **Excel.PivotFields** collection for a specific column (series), data, page, or row (category) field. The value is **-1** if the drop zone type is **xlDataField**.|
|PointIndex|Specifies the offset within the  **Office.IMsoSeries.Points** collection for a specific point within a series. A value of **?1** indicates that all data points are selected.|
|SeriesIndex|Specifies the offset within the  **Office.IMsoChart.SeriesCollection** for a specific series.|
|ShapeIndex|Specifies the offset within the [Shapes](http://msdn.microsoft.com/library/23aed165-e817-48b9-a7b8-050b81834494%28Office.15%29.aspx) collection for a specific shape.|
|TrendlineIndex|Specifies the offset within the  **Office.IMsoSeries.Trendlines** collection for a specific trendline within a series.|

## Example
<a name="pj15_VBAGetChartElement_Remarks"> </a>

The following example gets the chart element information for point (100, 100) in the chart. For example, if the point is within the plot area, output in the Immediate pane is  `idNum: 19, a: 0, b: 0`. From the information in Table 1,  **xlPlotArea** = 19.


```vb
Sub TestGetChartElements()
    Dim chartShape As Shape
    Dim reportName As String
    Dim x As Long
    Dim y As Long
    Dim idNum As Long
    Dim a As Long
    Dim b As Long
    
    reportName = "Simple scalar chart"
    Set chartShape = ActiveProject.Reports(reportName).Shapes(1)
    
    ' Specify a point in the chart.
    x = 100
    y = 100
    
    chartShape.Chart.GetChartElement x, y, idNum, a, b
    
    Debug.Print "idNum: " &; idNum &; ", a: " &; a &; ", b: " &; b
End Sub
```


## See also
<a name="pj15_VBAGetChartElement_Remarks"> </a>


#### Other resources


[Chart Object](chart-object-project.md)
