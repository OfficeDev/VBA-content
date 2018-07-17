---
title: LinearGradient.Degree Property (Excel)
keywords: vbaxl10.chm855074
f1_keywords:
- vbaxl10.chm855074
ms.prod: excel
api_name:
- Excel.LinearGradient.Degree
ms.assetid: 0608fe59-76e9-e199-2cc6-848f283813f3
ms.date: 06/08/2017
---


# LinearGradient.Degree Property (Excel)

The angle of the linear gradient fill within a selection. Read/write


## Syntax

 _expression_ . **Degree**

 _expression_ A variable that represents a **LinearGradient** object.


### Return Value

Double


## Remarks

Uses values ranging from 0 - 360.


## Example


```vb
With Selection.Interior 
 .Pattern = xlPatternLinearGradient 
 .Gradient.Degree = 45 
End With
```


## See also


#### Concepts


[LinearGradient Object](lineargradient-object-excel.md)

