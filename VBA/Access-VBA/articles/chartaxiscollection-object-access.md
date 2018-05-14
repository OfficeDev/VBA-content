---
title: ChartAxisCollection Object (Access)
keywords: vbaac10.chm14753
f1_keywords:
- vbaac10.chm14753
ms.prod: access
api_name:
- Access.ChartAxisCollection
ms.date: 05/02/2018
---


# ChartAxisCollection Object (Access)

A collection of all the **[ChartAxis](chartaxis-object.md)** objects in the specified chart.


## Using ChartAxisCollection

The following example displays the number of axes in the collection, then displays the name of each axis.

```vb
With myChart
 MsgBox (.ChartAxisCollection.Count)
  For Each axis In .ChartAxisCollection
    MsgBox (axis.Name)
  Next
End With
```

## See also


#### Concepts


[Chart Object](chart-object-access.md)