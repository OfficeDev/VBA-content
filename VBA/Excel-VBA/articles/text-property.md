---
title: Text Property
keywords: vbagr10.chm3077592
f1_keywords:
- vbagr10.chm3077592
ms.prod: excel
api_name:
- Excel.Text
ms.assetid: 1af6b778-b303-2bf1-8558-f665c71222a8
ms.date: 06/08/2017
---


# Text Property

Returns or sets the text for the specified object. Read/write String.

 _expression_. **Text**

 _expression_ Required. An expression that returns one of the objects in the Applies To list.


## Example

This example sets the text for the title of the chart.


```vb
With myChart 
 .HasTitle = True 
 .ChartTitle.Text = "First Quarter Sales" 
End With
```

This example sets the axis title text for the category axis.




```vb
With myChart.Axes(xlCategory) 
 .HasTitle = True 
 .AxisTitle.Text = "Month" 
End With
```


