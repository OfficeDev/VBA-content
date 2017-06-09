---
title: TickLabels.Offset Property (Word)
keywords: vbawd10.chm167051282
f1_keywords:
- vbawd10.chm167051282
ms.prod: word
api_name:
- Word.TickLabels.Offset
ms.assetid: f2435b6d-09a6-4dd9-eb51-71d7a1bf18c7
ms.date: 06/08/2017
---


# TickLabels.Offset Property (Word)

Returns or sets the distance between the levels of labels, and the distance between the first level and the axis line. Read/write  **Long** .


## Syntax

 _expression_ . **Offset**

 _expression_ A variable that represents a **[TickLabels](ticklabels-object-word.md)** object.


## Remarks

 The default distance is 100 percent, which represents the default spacing between the axis labels and the axis line. The value can be an integer percentage from 0 through 1000, relative to the axis label's font size.


## Example

The following example sets the label spacing of the category axis for the first chart in the active document to twice the current setting, if the offset is less than 500.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.Axes(xlCategory).TickLabels 
 If .Offset < 500 then 
 .Offset = .Offset * 2 
 End If 
 End With 
 End If 
End With
```


## See also


#### Concepts


[TickLabels Object](ticklabels-object-word.md)

