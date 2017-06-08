---
title: Chart.Perspective Property (Excel)
keywords: vbaxl10.chm149130
f1_keywords:
- vbaxl10.chm149130
ms.prod: excel
api_name:
- Excel.Chart.Perspective
ms.assetid: 39367c4a-95a7-afe7-b3e4-29e10a88fbd3
ms.date: 06/08/2017
---


# Chart.Perspective Property (Excel)

Returns or sets a  **Long** value that represents the perspective for the 3-D chart view.


## Syntax

 _expression_ . **Perspective**

 _expression_ A variable that represents a **Chart** object.


## Remarks

The value of this property must be between 0 and 100. This property is ignored if the  **[RightAngleAxes](chart-rightangleaxes-property-excel.md)** property is **True** .


## Example

This example sets the perspective of Chart1 to 70. The example should be run on a 3-D chart.


```vb
Charts("Chart1").RightAngleAxes = False 
Charts("Chart1").Perspective = 70
```


## See also


#### Concepts


[Chart Object](chart-object-excel.md)

