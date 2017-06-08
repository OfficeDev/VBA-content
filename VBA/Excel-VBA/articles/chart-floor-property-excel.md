---
title: Chart.Floor Property (Excel)
keywords: vbaxl10.chm149109
f1_keywords:
- vbaxl10.chm149109
ms.prod: excel
api_name:
- Excel.Chart.Floor
ms.assetid: 7771ab49-b254-f0f0-a21b-596f541ab6c1
ms.date: 06/08/2017
---


# Chart.Floor Property (Excel)

Returns a  **[Floor](floor-object-excel.md)** object that represents the floor of the 3-D chart. Read-only.


## Syntax

 _expression_ . **Floor**

 _expression_ An expression that returns a **Chart** object.


### Return Value

Floor


## Example

This example sets the floor color of Chart1 to blue. The example should be run on a 3-D chart (the  **Floor** property fails on 2-D charts).


```vb
Charts("Chart1").Floor.Interior.ColorIndex = 5
```


## See also


#### Concepts


[Chart Object](chart-object-excel.md)

