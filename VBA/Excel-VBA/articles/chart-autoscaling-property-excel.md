---
title: Chart.AutoScaling Property (Excel)
keywords: vbaxl10.chm149080
f1_keywords:
- vbaxl10.chm149080
ms.prod: excel
api_name:
- Excel.Chart.AutoScaling
ms.assetid: fecafb42-56fb-3c33-dc03-cb290b4a28df
ms.date: 06/08/2017
---


# Chart.AutoScaling Property (Excel)

 **True** if Microsoft Excel scales a 3-D chart so that it's closer in size to the equivalent 2-D chart. The **[RightAngleAxes](chart-rightangleaxes-property-excel.md)** property must be **True** . Read/write **Boolean** .


## Syntax

 _expression_ . **AutoScaling**

 _expression_ A variable that represents a **Chart** object.


## Example

This example automatically scales Chart1. The example should be run on a 3-D chart.


```vb
With Charts("Chart1") 
 .RightAngleAxes = True 
 .AutoScaling = True 
End With
```


## See also


#### Concepts


[Chart Object](chart-object-excel.md)

