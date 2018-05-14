---
title: Chart.ChartValues Property (Access)
keywords: vbaac10.chm6108
f1_keywords:
- vbaac10.chm6108
ms.prod: access
api_name:
- Access.Chart.ChartValues
ms.date: 05/02/2018
---


# Chart.ChartValues Property (Access)

Returns or sets the semicolon-separated list of field(s) used to determine the data series plotted on the value axis. Read/write **String** .


## Syntax

 _expression_ . **ChartValues**

 _expression_ A variable that represents a **Chart** object.


## Example

```vb
With myChart
 .ChartValues = "[Price];[Cost]"
End With
```

## See also


#### Concepts


[Chart Object](chart-object-access.md)