---
title: Series.LeaderLines Property (Word)
keywords: vbawd10.chm123733634
f1_keywords:
- vbawd10.chm123733634
ms.prod: word
api_name:
- Word.Series.LeaderLines
ms.assetid: 5b4f8802-2b1f-a879-f74d-b98a82ba9187
ms.date: 06/08/2017
---


# Series.LeaderLines Property (Word)

Returns the leader lines for the series. Read-only  **[LeaderLines](leaderlines-object-word.md)** .


## Syntax

 _expression_ . **LeaderLines**

 _expression_ A variable that represents a **[Series](series-object-word.md)** object.


## Remarks

This property applies only to pie charts.


## Example

The following example adds data labels and blue leader lines to series one on the first pie chart in the active document. If no leader lines are visible, this example code will fail. In this situation, you can manually drag one of the data labels away from the pie chart to make a leader line show up.


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

