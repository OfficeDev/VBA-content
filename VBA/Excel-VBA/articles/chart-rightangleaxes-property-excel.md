---
title: Chart.RightAngleAxes Property (Excel)
keywords: vbaxl10.chm149138
f1_keywords:
- vbaxl10.chm149138
ms.prod: excel
api_name:
- Excel.Chart.RightAngleAxes
ms.assetid: 632aa454-4113-97d3-a80c-eb745a950c6f
ms.date: 06/08/2017
---


# Chart.RightAngleAxes Property (Excel)

 **True** if the chart axes are at right angles, independent of chart rotation or elevation. Applies only to 3-D line, column, and bar charts. Read/write **Boolean** .


## Syntax

 _expression_ . **RightAngleAxes**

 _expression_ A variable that represents a **Chart** object.


## Remarks

If this property is  **True** , the **[Perspective](chart-perspective-property-excel.md)** property is ignored.


## Example

This example sets the axes in Chart1 to intersect at right angles. The example should be run on a 3-D chart.


```vb
Charts("Chart1").RightAngleAxes = True
```


## See also


#### Concepts


[Chart Object](chart-object-excel.md)

