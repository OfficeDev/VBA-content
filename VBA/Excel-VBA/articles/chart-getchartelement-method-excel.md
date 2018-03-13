---
title: Chart.GetChartElement Method (Excel)
keywords: vbaxl10.chm149161
f1_keywords:
- vbaxl10.chm149161
ms.prod: excel
api_name:
- Excel.Chart.GetChartElement
ms.assetid: a4888d1b-f73b-43cd-5318-95c1d63944fa
ms.date: 06/08/2017
---


# Chart.GetChartElement Method (Excel)

Returns information about the chart element at specified X and Y coordinates. This method is unusual in that you specify values for only the first two arguments. Microsoft Excel fills in the other arguments, and your code should examine those values when the method returns.


## Syntax

 _expression_ . **GetChartElement**( **_x_** , **_y_** , **_ElementID_** , **_Arg1_** , **_Arg2_** )

 _expression_ A variable that represents a **Chart** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _x_|Required| **Long**|The X coordinate of the chart element.|
| _y_|Required| **Long**|The Y coordinate of the chart element.|
| _ElementID_|Required| **Long**|When the method returns, this argument contains the  **[XLChartItem](xlchartitem-enumeration-excel.md)** value of the chart element at the specified coordinates. For more information, see the ?Remarks? section.|
| _Arg1_|Required| **Long**|When the method returns, this argument contains information related to the chart element. For more information, see the ?Remarks? section.|
| _Arg2_|Required| **Long**|When the method returns, this argument contains information related to the chart element. For more information, see the ?Remarks? section.|

## Remarks

The value of  _ElementID_ after the method returns determines whether _Arg1_ and _Arg2_ contain any information, as shown in the following table.



| <strong>ElementID Constant</strong>      | **Constant Value ** | <strong>Arg1</strong> | <strong>Arg2</strong> |
|:-----------------------------------------|:--------------------|:----------------------|:----------------------|
| <strong>xlAxis</strong>                  | 21                  | AxisIndex             | AxisType              |
| <strong>xlAxisTitle</strong>             | 17                  | AxisIndex             | AxisType              |
| <strong>xlDisplayUnitLabel</strong>      | 30                  | AxisIndex             | AxisType              |
| <strong>xlMajorGridlines</strong>        | 15                  | AxisIndex             | AxisType              |
| <strong>xlMinorGridlines</strong>        | 16                  | AxisIndex             | AxisType              |
| <strong>xlPivotChartDropZone</strong>    | 32                  | DropZoneType          | None                  |
| <strong>xlPivotChartFieldButton</strong> | 31                  | DropZoneType          | PivotFieldIndex       |
| <strong>xlDownBars</strong>              | 20                  | GroupIndex            | None                  |
| <strong>xlDropLines</strong>             | 26                  | GroupIndex            | None                  |
| <strong>xlHiLoLines</strong>             | 25                  | GroupIndex            | None                  |
| <strong>xlRadarAxisLabels</strong>       | 27                  | GroupIndex            | None                  |
| <strong>xlSeriesLines</strong>           | 22                  | GroupIndex            | None                  |
| <strong>xlUpBars</strong>                | 18                  | GroupIndex            | None                  |
| <strong>xlChartArea</strong>             | 2                   | None                  | None                  |
| <strong>xlChartTitle</strong>            | 4                   | None                  | None                  |
| <strong>xlCorners</strong>               | 6                   | None                  | None                  |
| <strong>xlDataTable</strong>             | 7                   | None                  | None                  |
| <strong>xlFloor</strong>                 | 23                  | None                  | None                  |
| <strong>xlLeaderLines</strong>           | 29                  | None                  | None                  |
| <strong>xlLegend</strong>                | 24                  | None                  | None                  |
| <strong>xlNothing</strong>               | 28                  | None                  | None                  |
| <strong>xlPlotArea</strong>              | 19                  | None                  | None                  |
| <strong>xlWalls</strong>                 | 5                   | None                  | None                  |
| <strong>xlDataLabel</strong>             | 7                   | SeriesIndex           | PointIndex            |
| <strong>xlErrorBars</strong>             | 9                   | SeriesIndex           | None                  |
| <strong>xlLegendEntry</strong>           | 12                  | SeriesIndex           | None                  |
| <strong>xlLegendKey</strong>             | 13                  | SeriesIndex           | None                  |
| <strong>xlSeries</strong>                | 3                   | SeriesIndex           | PointIndex            |
| <strong>xlShape</strong>                 | 14                  | ShapeIndex            | None                  |
| <strong>xlTrendline</strong>             | 8                   | SeriesIndex           | TrendLineIndex        |
| <strong>xlXErrorBars</strong>            | 10                  | SeriesIndex           | None                  |
| <strong>xlYErrorBars</strong>            | 11                  | SeriesIndex           | None                  |

The following table describes the meaning of  _Arg1_ and _Arg2_ after the method returns.



|**Argument**|**Description**|
|:-----|:-----|
|AxisIndex|Specifies whether the axis is primary or secondary. Can be one of the following  **[XlAxisGroup](xlaxisgroup-enumeration-excel.md)** constants: **xlPrimary** or **xlSecondary** .|
|AxisType|Specifies the axis type. Can be one of the following  **[XlAxisType](xlaxistype-enumeration-excel.md)** constants: **xlCategory** , **xlSeriesAxis** , or **xlValue** .|
|DropZoneType|Specifies the drop zone type: column, data, page, or row field. Can be one of the following  **[XlPivotFieldOrientation](xlpivotfieldorientation-enumeration-excel.md)** constants: **xlColumnField** , **xlDataField** , **xlPageField** , or **xlRowField** . The column and row field constants specify the series and category fields, respectively.|
|GroupIndex|Specifies the offset within the  **[ChartGroups](chartgroups-object-excel.md)** collection for a specific chart group.|
|PivotFieldIndex|Specifies the offset within the  **[PivotFields](pivotfields-object-excel.md)** collection for a specific column (series), data, page, or row (category) field. -1 if the drop zone type is **xlDataField** .|
|PointIndex|Specifies the offset within the  **[Points](points-object-excel.md)** collection for a specific point within a series. A value of ? 1 indicates that all data points are selected.|
|SeriesIndex|Specifies the offset within the  **[Series](series-object-excel.md)** collection for a specific series.|
|ShapeIndex|Specifies the offset within the  **[Shapes](shapes-object-excel.md)** collection for a specific shape.|
|TrendlineIndex|Specifies the offset within the  **[Trendlines](trendlines-object-excel.md)** collection for a specific trendline within a series.|

## Example

This example warns the user if she moves the mouse over the chart legend.


```vb
Private Sub Chart_MouseMove(ByVal Button As Long, _ 
 ByVal Shift As Long, ByVal X As Long, ByVal Y As Long) 
 Dim IDNum As Long 
 Dim a As Long 
 Dim b As Long 

 ActiveChart.GetChartElement X, Y, IDNum, a, b 
 If IDNum = xlLegendEntry Then _ 
 MsgBox "WARNING: Move away from the legend" 
End Sub
```


## See also


#### Concepts


[Chart Object](chart-object-excel.md)

