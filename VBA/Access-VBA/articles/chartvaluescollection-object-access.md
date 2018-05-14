---
title: ChartValuesCollection Object (Access)
keywords: vbaac10.chm14755
f1_keywords:
- vbaac10.chm14755
ms.prod: access
api_name:
- Access.ChartValuesCollection
ms.date: 05/02/2018
---


# ChartValuesCollection Object (Access)

A collection of all the **[ChartValues](chartvalues-object.md)** objects in the specified chart.


## Using ChartValuesCollection

The following example displays the name of each **[ChartValues](chartvalues-object.md)** instance in a collection.

```vb
With myChart
 For Each cv In .ChartValuesCollection
  MsgBox (cv.Name)
 Next
End With
```

## See also


#### Concepts


[Chart Object](chart-object-access.md)