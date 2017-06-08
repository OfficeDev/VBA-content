---
title: Chart.ChartTitle Property (Excel)
keywords: vbaxl10.chm149089
f1_keywords:
- vbaxl10.chm149089
ms.prod: excel
api_name:
- Excel.Chart.ChartTitle
ms.assetid: 3a083c1f-7a3f-3368-c547-297f0e5d26cb
ms.date: 06/08/2017
---


# Chart.ChartTitle Property (Excel)

Returns a  **[ChartTitle](charttitle-object-excel.md)** object that represents the title of the specified chart. Read-only.


## Syntax

 _expression_ . **ChartTitle**

 _expression_ A variable that represents a **Chart** object.


## Example

This example sets the text for the title of Chart1.


```vb
With Charts("Chart1") 
 .HasTitle = True 
 .ChartTitle.Text = "First Quarter Sales" 
End With
```


## See also


#### Concepts


[Chart Object](chart-object-excel.md)

