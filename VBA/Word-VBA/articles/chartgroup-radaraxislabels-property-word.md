---
title: ChartGroup.RadarAxisLabels Property (Word)
keywords: vbawd10.chm263454744
f1_keywords:
- vbawd10.chm263454744
ms.prod: word
api_name:
- Word.ChartGroup.RadarAxisLabels
ms.assetid: 30b37487-bef9-b333-7df7-546d85a92047
ms.date: 06/08/2017
---


# ChartGroup.RadarAxisLabels Property (Word)

Returns the radar axis labels for the specified chart group. Read-only  **[TickLabels](ticklabels-object-word.md)** .


## Syntax

 _expression_ . **RadarAxisLabels**

 _expression_ A variable that represents a **[ChartGroup](chartgroup-object-word.md)** object.


## Example

The following example enables radar axis labels for chart group one for the first chart in the active document and then sets the color for the labels to red. You should run the example on a radar chart.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.ChartGroups(1) 
 .HasRadarAxisLabels = True 
 .RadarAxisLabels.Font.ColorIndex = 3 
 End With 
 End If 
End With
```


## See also


#### Concepts


[ChartGroup Object](chartgroup-object-word.md)

