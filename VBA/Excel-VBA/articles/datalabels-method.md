---
title: DataLabels Method
keywords: vbagr10.chm3077616
f1_keywords:
- vbagr10.chm3077616
ms.prod: excel
api_name:
- Excel.DataLabels
ms.assetid: 8ffca32c-f505-482e-dd27-d29ad2682daf
ms.date: 06/08/2017
---


# DataLabels Method

Returns an object that represents either a single data label or a collection of all the data labels for the series.

 _expression_. **DataLabels**( **_Index_**)

 _expression_ Required. An expression that returns one of the objects in the Applies To list.

 **Index** Optional **Variant**. The number of the data label.

## Example

This example sets the data labels for series one to show their key, assuming that their values are visible when the example runs.


```vb
With myChart.SeriesCollection(1) 
 .HasDataLabels = True 
 With .DataLabels 
 .ShowLegendKey = True 
 .Type = xlValue 
 End With 
End With
```


