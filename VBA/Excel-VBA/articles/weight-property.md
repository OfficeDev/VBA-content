---
title: Weight Property
keywords: vbagr10.chm65656
f1_keywords:
- vbagr10.chm65656
ms.prod: excel
api_name:
- Excel.Weight
ms.assetid: 59a3b106-5811-f082-d9cf-c21f2945da31
ms.date: 06/08/2017
---


# Weight Property

Returns or sets the weight of the border. Read/write 
 **XlBorderWeight**
.



|XlBorderWeight can be one of these XlBorderWeight constants.|
| **xlHairline**|
| **xlThin** **xlMedium** **xlThick**|

 _expression_. **Weight**

 _expression_ Required. An expression that returns one of the objects in the Applies To list.

## Example

This example sets the border weight for the chart area.


```
myChart.ChartArea.Border.Weight = xlMedium
```


