---
title: AutoScaling Property
keywords: vbagr10.chm65643
f1_keywords:
- vbagr10.chm65643
ms.prod: excel
api_name:
- Excel.AutoScaling
ms.assetid: f132291c-e356-eea5-0ef5-0e4def8d4832
ms.date: 06/08/2017
---


# AutoScaling Property

True if Microsoft Graph scales a 3-D chart so that it's closer in size to the equivalent 2-D chart. The RightAngleAxes property must be True. Read/write Boolean.

 _expression_. **AutoScaling**

 _expression_ Required. An expression that returns one of the objects in the Applies To list.


## Example

This example automatically scales the chart. The example should be run on a 3-D chart.


```vb
With myChart 
 .RightAngleAxes = True 
 .AutoScaling = True 
End With
```


