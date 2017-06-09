---
title: Clear Method
keywords: vbagr10.chm65647
f1_keywords:
- vbagr10.chm65647
ms.prod: excel
api_name:
- Excel.Clear
ms.assetid: f77c2fc0-6ec4-7345-0e5c-7b8dd4cd1a90
ms.date: 06/08/2017
---


# Clear Method

Clear method as it applies to the  **ChartArea** and **Legend** objects.

Clears the entire chart area.

 _expression_. **Clear**

 _expression_ Required. An expression that returns one of the above objects.
Clear method as it applies to the  **Range** object.
Clears the entire range.
 _expression_. **Clear**
 _expression_ Required. An expression that returns one of the above objects.

## Example

This example clears the formulas and formatting in cells A1:G37 on the datasheet.


```
myChart.Application.DataSheet.Range("A1:G37").Clear
```

This example clears the chart area (the chart data and formatting) of Chart1.




```
myChart.ChartArea.Clear
```


