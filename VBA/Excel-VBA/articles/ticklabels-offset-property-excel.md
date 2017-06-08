---
title: TickLabels.Offset Property (Excel)
keywords: vbaxl10.chm616085
f1_keywords:
- vbaxl10.chm616085
ms.prod: excel
api_name:
- Excel.TickLabels.Offset
ms.assetid: a353b803-34a3-0ff9-83d2-3318c308ec35
ms.date: 06/08/2017
---


# TickLabels.Offset Property (Excel)

Returns or sets a  **Long** value that represents the distance between the levels of labels, and the distance between the first level and the axis line.


## Syntax

 _expression_ . **Offset**

 _expression_ A variable that represents a **TickLabels** object.


## Remarks

 The default distance is 100 percent, which represents the default spacing between the axis labels and the axis line. The value can be an integer percentage from 0 through 1000, relative to the axis label's font size.


## Example

This example sets the label spacing of the category axis in Chart1 to twice the current setting, if the offset is less than 500.


```vb
With Charts("Chart1").Axes(xlCategory).TickLabels 
 If .Offset < 500 then 
 .Offset = .Offset * 2 
 End If 
End With
```


## See also


#### Concepts


[TickLabels Object](ticklabels-object-excel.md)

