---
title: Interior Object
keywords: vbagr10.chm5207570
f1_keywords:
- vbagr10.chm5207570
ms.prod: excel
api_name:
- Excel.Interior
ms.assetid: 13a4801e-f121-2a43-cd61-cf3ac9325197
ms.date: 06/08/2017
---


# Interior Object

Represents the interior of the specified object.


## Using the Interior Object

Use the  **Interior** property to return the **Interior** object. The following example sets the chart area color to gray and the plot area color to green.


```vb
With myChart 
 .PlotArea.Interior.Color = RGB(0, 100, 150) 
 .ChartArea.Interior.Color = RGB(50, 10, 50) 
End With
```


