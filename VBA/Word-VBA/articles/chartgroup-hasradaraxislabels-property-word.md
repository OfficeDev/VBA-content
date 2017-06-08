---
title: ChartGroup.HasRadarAxisLabels Property (Word)
keywords: vbawd10.chm263454734
f1_keywords:
- vbawd10.chm263454734
ms.prod: word
api_name:
- Word.ChartGroup.HasRadarAxisLabels
ms.assetid: 0b086c3c-1eaa-1e65-fcb1-969c8b2c64c7
ms.date: 06/08/2017
---


# ChartGroup.HasRadarAxisLabels Property (Word)

 **True** if a radar chart has axis labels. Read/write **Boolean** .


## Syntax

 _expression_ . **HasRadarAxisLabels**

 _expression_ A variable that represents a **[ChartGroup](chartgroup-object-word.md)** object.


## Remarks

This property applies only to radar charts. 


## Example

The following example enables radar axis labels for chart group one of the first chart in the active document and sets their color. You should run the example on a radar chart.


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

