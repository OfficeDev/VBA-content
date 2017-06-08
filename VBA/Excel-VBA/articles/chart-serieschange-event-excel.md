---
title: Chart.SeriesChange Event (Excel)
keywords: vbaxl10.chm500084
f1_keywords:
- vbaxl10.chm500084
ms.prod: excel
api_name:
- Excel.Chart.SeriesChange
ms.assetid: 80a8058c-0445-0051-24d1-1a965c302790
ms.date: 06/08/2017
---


# Chart.SeriesChange Event (Excel)

Occurs when the user changes the value of a chart data point by clicking a bar in the chart and dragging the top edge up or down thus changing the value of the data point.


 **Important**  This event is not functional in Excel 2007 and later versions. You should not use it in your code.


## Syntax

 _expression_ . **SeriesChange**( **_SeriesIndex_** , **_PointIndex_** )

 _expression_ A variable that represents a **Chart** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SeriesIndex_|Required| **Long**| The offset within the **[Series](series-object-excel.md)** collection for the changed series.|
| _PointIndex_|Required| **Long**|The offset within the  **[Points](points-object-excel.md)** collection for the changed point.|

### Return Value

Nothing


## Example

This example changes the point's border color when the user changes the point value.


```vb
Private Sub Chart_SeriesChange(ByVal SeriesIndex As Long, _ 
 ByVal PointIndex As Long) 
 Set p = Me.SeriesCollection(SeriesIndex).Points(PointIndex) 
 p.Border.ColorIndex = 3 
End Sub
```


## See also


#### Concepts


[Chart Object](chart-object-excel.md)

