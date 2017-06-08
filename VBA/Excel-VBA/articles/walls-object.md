---
title: Walls Object
keywords: vbagr10.chm5208137
f1_keywords:
- vbagr10.chm5208137
ms.prod: excel
api_name:
- Excel.Walls
ms.assetid: 97c3a312-abf1-9da7-cbff-8e48737bf499
ms.date: 06/08/2017
---


# Walls Object

Represents the walls of the specified 3-D chart. This object isn't a collection. There's no object that represents a single wall; you must return all the walls as a unit.


## Using the Walls Object

Use the  **Walls** property to return the **Walls** object. The following example sets the pattern on the walls for the chart. If the chart isn't a 3-D chart, this example will fail.


```
myChart.Walls.Interior.Pattern = xlGray75
```


