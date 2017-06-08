---
title: Chart.ChartArea Property (Excel)
keywords: vbaxl10.chm149086
f1_keywords:
- vbaxl10.chm149086
ms.prod: excel
api_name:
- Excel.Chart.ChartArea
ms.assetid: 125d6176-b770-900b-8572-ce33b95ad897
ms.date: 06/08/2017
---


# Chart.ChartArea Property (Excel)

Returns a  **[ChartArea](chartarea-object-excel.md)** object that represents the complete chart area for the chart. Read-only.


## Syntax

 _expression_ . **ChartArea**

 _expression_ A variable that represents a **Chart** object.


## Example

This example sets the chart area interior color of Chart1 to red and sets the border color to blue.


```vb
With Charts("Chart1").ChartArea 
 .Interior.ColorIndex = 3 
 .Border.ColorIndex = 5 
End With
```


## See also


#### Concepts


[Chart Object](chart-object-excel.md)

