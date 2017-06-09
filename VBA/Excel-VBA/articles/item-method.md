---
title: Item Method
keywords: vbagr10.chm3077621
f1_keywords:
- vbagr10.chm3077621
ms.prod: excel
api_name:
- Excel.Item
ms.assetid: 9e92de7f-b231-c7c5-fcea-50c1051d1add
ms.date: 06/08/2017
---


# Item Method

Item method as it applies to the  **Axes** object.

Returns a single Axis object from an Axes collection.

 _expression_. **Item**( **_Type_**,  **_AxisGroup_**)

 _expression_ Required. An expression that returns an **Axes** collection.
 **Type**Required 
 **XlAxisType**
. The axis type.


|XlAxisType can be one of these XlAxisType constants.|
| **xlCategory**|
| **xlSeriesAxis** Valid only for 3-D charts.|
| **xlValue**|
 **AxisGroup**Optional 
 **XlAxisGroup**
. The axis group.


|XlAxisGroup can be one of these XlAxisGroup constants.|
| **xlSecondary**|
| **xlPrimary**_default_|
Item method as it applies to the  **ChartGroups** object.
Returns a single ChartGroup object from a ChartGroups collection.
 _expression_. **Item**( **_Index_**)
 _expression_ Required. An expression that returns a **ChartGroups** collection.
 **Index**Required  **Variant**. The index number of the chart group.
Item method as it applies to the  **DataLabels** object.
Returns a single DataLabel object from a DataLabels collection.
 _expression_. **Item**( **_Index_**)
 _expression_ Required. An expression that returns a **DataLabels** collection.
 **Index**Required  **Variant**. The name or index number of the data label.
Item method as it applies to the  **LegendEntries** object.
Returns a single LegendEntry object from a LegendEntries collection.
 _expression_. **Item**( **_Index_**)
 _expression_ Required. An expression that returns a **LegendEntries** collection.
 **Index**Required  **Variant**. The index number of the legend entry.
Item method as it applies to the  **Points** object.
Returns a single Point object from a Points collection.
 _expression_. **Item**( **_Index_**)
 _expression_ Required. An expression that returns a **Points** collection.
 **Index**Required  **Long**. The index number of the point.
Item method as it applies to the  **SeriesCollection** object.
Returns a single Series object from a SeriesCollection collection.
 _expression_. **Item**( **_Index_**)
 _expression_ Required. An expression that returns a **SeriesCollection** collection.
 **Index**Required  **Variant**. The name or index number of the series.
Item method as it applies to the  **Trendlines** object.
Returns a single Trendline object from a Trendlines collection.
 _expression_. **Item**( **_Index_**)
 _expression_ Required. An expression that returns a **Trendlines** collection.
 **Index**Optional  **Variant**. The name or index number of the trendline.

## Example

As it applies to the  **Axes** object.

This example sets the title text for the category axis on Chart1.




```vb
With Charts("chart1").Axes.Item(xlCategory) 
 .HasTitle = True 
 .AxisTitle.Caption = "1994" 
End With
```

As it applies to the  **ChartGroups** object.

This example adds drop lines to chart group one on chart sheet one.




```vb
Charts(1).ChartGroups.Item(1).HasDropLines = True
```

As it applies to the  **DataLabels** object.

This example sets the number format for the fifth data label in series one in embedded chart one on worksheet one.




```vb
Worksheets(1).ChartObjects(1).Chart _ 
 .SeriesCollection(1).DataLabels.Item(5).NumberFormat = "0.000"
```

As it applies to the  **LegendEntries** object.

This example changes the font for the text of the legend entry at the top of the legend (this is usually the legend for series one) in embedded chart one on Sheet1.




```vb
Worksheets("sheet1").ChartObjects(1).Chart _ 
 .Legend.LegendEntries.Item(1).Font.Italic = True
```

As it applies to the  **Points** object.

This example sets the marker style for the third point in series one in embedded chart one on worksheet one. The specified series must be a 2-D line, scatter, or radar series.




```vb
Worksheets(1).ChartObjects(1).Chart. _ 
 SeriesCollection(1).Points.Item(3).MarkerStyle = xlDiamond
```

As it applies to the  **SeriesCollection** object.

This example provides two lines of code that are equivalent:




```
myChart.SeriesCollection.Item(1) 
myChart.SeriesCollection(1)
```

As it applies to the  **Trendlines** object.

This example sets the number of units that the trendline on Chart1 extends forward and backward. The example should be run on a 2-D column chart that contains a single series with a trendline.




```vb
With Charts("Chart1").SeriesCollection(1).Trendlines.Item(1) 
 .Forward = 5 
 .Backward = .5 
End With
```


