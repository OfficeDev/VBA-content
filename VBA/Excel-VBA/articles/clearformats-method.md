---
title: ClearFormats Method
keywords: vbagr10.chm3077613
f1_keywords:
- vbagr10.chm3077613
ms.prod: excel
api_name:
- Excel.ClearFormats
ms.assetid: a238ae6f-a673-f49b-1bd5-414d93beb97e
ms.date: 06/08/2017
---


# ClearFormats Method

Clears the formatting of the object.

 _expression_. **ClearFormats**

 _expression_ Required. An expression that returns one of the objects in the Applies To list.


## Example

This example clears all formatting from cells A1:G37 on the datasheet.


```
myChart.Application.DataSheet.Range("A1:G37").ClearFormats
```

This example clears the formatting from the chart.




```
myChart.ChartArea.ClearFormats
```


