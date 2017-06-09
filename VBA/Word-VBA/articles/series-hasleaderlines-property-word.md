---
title: Series.HasLeaderLines Property (Word)
keywords: vbawd10.chm123733362
f1_keywords:
- vbawd10.chm123733362
ms.prod: word
api_name:
- Word.Series.HasLeaderLines
ms.assetid: c558ffc3-939b-a237-3c6e-e10549f3c8d8
ms.date: 06/08/2017
---


# Series.HasLeaderLines Property (Word)

 **True** if the series has leader lines. Read/write **Boolean** .


## Syntax

 _expression_ . **HasLeaderLines**

 _expression_ A variable that represents a **[Series](series-object-word.md)** object.


## Remarks

This property applies only to pie charts.


## Example

The following example adds data labels and blue leader lines to series one on the pie chart. If no leader lines are visible, this example code will fail. In this situation, you can manually drag one of the data labels away from the pie chart to make a leader line show up.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.SeriesCollection(1) 
 .HasDataLabels = True 
 .DataLabels.Position = xlLabelPositionBestFit 
 .HasLeaderLines = True 
 .LeaderLines.Border.ColorIndex = 5 
 End With 
 End If 
End With 

```


## See also


#### Concepts


[Series Object](series-object-word.md)

