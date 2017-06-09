---
title: SparklineGroups.Ungroup Method (Excel)
keywords: vbaxl10.chm869081
f1_keywords:
- vbaxl10.chm869081
ms.prod: excel
api_name:
- Excel.SparklineGroups.Ungroup
ms.assetid: c67c54f4-d5d1-5f12-2413-671db612a954
ms.date: 06/08/2017
---


# SparklineGroups.Ungroup Method (Excel)

Ungroups the sparklines in the selected sparkline group.


## Syntax

 _expression_ . **Ungroup**

 _expression_ A variable that represents a **[SparklineGroups](sparklinegroups-object-excel.md)** object.


### Return Value

Nothing


## Example

The following code example selects the range A1:A4 and ungroups the sparklines in that range.


```vb
Range("A1:A4").Select 
Selection.SparklineGroups.Ungroup
```


## See also


#### Concepts


[SparklineGroups Object](sparklinegroups-object-excel.md)

