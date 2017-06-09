---
title: Chart.BeforeDoubleClick Event (Excel)
keywords: vbaxl10.chm500082
f1_keywords:
- vbaxl10.chm500082
ms.prod: excel
api_name:
- Excel.Chart.BeforeDoubleClick
ms.assetid: 406c6b9f-1182-5f5b-b954-afe10cd21a9b
ms.date: 06/08/2017
---


# Chart.BeforeDoubleClick Event (Excel)

Occurs when a chart element is double-clicked, before the default double-click action.


## Syntax

 _expression_ . **BeforeDoubleClick**( **_ElementID_** , **_Arg1_** , **_Arg2_** , **_Cancel_** )

 _expression_ A variable that represents a **Chart** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True** , the default double-click action isn't performed when the procedure is finished.|
| _Arg1_|Required| **Long**|Additional event information, depending on the value of  _ElementID_. For more information about this parameter, see the Remarks section.|
| _Arg2_|Required| **Long**|Additional event information, depending on the value of  _ElementID_. For more information about this parameter, see the Remarks section.|
| _ElementID_|Required| **Long**|The double-clicked object. The value of this parameter determines the expected values of  _Arg1_ and _Arg2_. For more information about this paramter, see the Remarks section.|

## Remarks

The  **[DoubleClick](application-doubleclick-method-excel.md)** method doesn't cause this event to occur.

This event doesn't occur when the user double-clicks the border of a cell.

The meaning of  _Arg1_ and _Arg2_ depends on the _ElementID_ value, as shown in the following table.



|**_ElementID_**|**_Arg1_**|**_Arg2_**|
|:-----|:-----|:-----|
| **xlAxis**|AxisIndex|AxisType|
| **xlAxisTitle**|AxisIndex|AxisType|
| **xlDisplayUnitLabel**|AxisIndex|AxisType|
| **xlMajorGridlines**|AxisIndex|AxisType|
| **xlMinorGridlines**|AxisIndex|AxisType|
| **xlPivotChartDropZone**|DropZoneType|None|
| **xlPivotChartFieldButton**|DropZoneType|PivotFieldIndex|
| **xlDownBars**|GroupIndex|None|
| **xlDropLines**|GroupIndex|None|
| **xlHiLoLines**|GroupIndex|None|
| **xlRadarAxisLabels**|GroupIndex|None|
| **xlSeriesLines**|GroupIndex|None|
| **xlUpBars**|GroupIndex|None|
| **xlChartArea**|None|None|
| **xlChartTitle**|None|None|
| **xlCorners**|None|None|
| **xlDataTable**|None|None|
| **xlFloor**|None|None|
| **xlLegend**|None|None|
| **xlNothing**|None|None|
| **xlPlotArea**|None|None|
| **xlWalls**|None|None|
| **xlDataLabel**|SeriesIndex|PointIndex|
| **xlErrorBars**|SeriesIndex|None|
| **xlLegendEntry**|SeriesIndex|None|
| **xlLegendKey**|SeriesIndex|None|
| **xlSeries**|SeriesIndex|PointIndex|
| **xlTrendline**|SeriesIndex|TrendLineIndex|
| **xlXErrorBars**|SeriesIndex|None|
| **xlYErrorBars**|SeriesIndex|None|
| **xlShape**|ShapeIndex|None|
The following table describes the meaning of the arguments.



|**Argument**|**Description**|
|:-----|:-----|
|AxisIndex|Specifies whether the axis is primary or secondary. Can be one of the following  **XlAxisGroup** constants: **xlPrimary** or **xlSecondary** .|
|AxisType|Specifies the axis type. Can be one of the following  **XlAxisType** constants: **xlCategory** , **xlSeriesAxis** , or **xlValue** .|
|DropZoneType|Specifies the drop zone type: column, data, page, or row field. Can be one of the following  **XlPivotFieldOrientation** constants: **xlColumnField** , **xlDataField** , **xlPageField** , or **xlRowField** . The column and row field constants specify the series and category fields, respectively.|
|GroupIndex|Specifies the offset within the  **[ChartGroups](chartgroups-object-excel.md)** collection for a specific chart group.|
|PivotFieldIndex|Specifies the offset within the  **[PivotFields](pivotfields-object-excel.md)** collection for a specific column (series), data, page, or row (category) field.|
|PointIndex|Specifies the offset within the  **[Points](points-object-excel.md)** collection for a specific point within a series. The value ? 1 indicates that all data points are selected.|
|SeriesIndex|Specifies the offset within the  **[Series](series-object-excel.md)** collection for a specific series.|
|ShapeIndex|Specifies the offset within the  **[Shapes](shapes-object-excel.md)** collection for a specific shape.|
|TrendlineIndex|Specifies the offset within the  **[Trendlines](trendlines-object-excel.md)** collection for a specific trendline within a series.|

## Example

This example overrides the default double-click behavior for the chart floor.


```vb
Private Sub Chart_BeforeDoubleClick(ByVal ElementID As Long, _ 
 ByVal Arg1 As Long, ByVal Arg2 As Long, Cancel As Boolean) 
 
 If ElementID = xlFloor Then 
 Cancel = True 
 MsgBox "Chart formatting for this item is restricted." 
 End If 
 
End Sub
```


## See also


#### Concepts


[Chart Object](chart-object-excel.md)

