---
title: Chart.GapDepth Property (Excel)
keywords: vbaxl10.chm149110
f1_keywords:
- vbaxl10.chm149110
ms.prod: excel
api_name:
- Excel.Chart.GapDepth
ms.assetid: 6020490a-1343-5b79-ff7d-197f78061420
ms.date: 06/08/2017
---


# Chart.GapDepth Property (Excel)

Returns or sets the distance between the data series in a 3-D chart, as a percentage of the marker width. The value of this property must be between 0 and 500. Read/write  **Long** .


## Syntax

 _expression_ . **GapDepth**

 _expression_ A variable that represents a **Chart** object.


## Example

This example sets the distance between the data series in Chart1 to 200 percent of the marker width. The example should be run on a 3-D chart (the  **GapDepth** property fails on 2-D charts).


```vb
Charts("Chart1").GapDepth = 200
```


## See also


#### Concepts


[Chart Object](chart-object-excel.md)

