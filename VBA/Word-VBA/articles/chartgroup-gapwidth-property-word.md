---
title: ChartGroup.GapWidth Property (Word)
keywords: vbawd10.chm263454728
f1_keywords:
- vbawd10.chm263454728
ms.prod: word
api_name:
- Word.ChartGroup.GapWidth
ms.assetid: 7f8d7f6b-9086-19c2-c4f4-d947491631ec
ms.date: 06/08/2017
---


# ChartGroup.GapWidth Property (Word)

For bar and column charts, returns or sets the space, as a percentage of the bar or column width, between bar or column clusters. For pie-of-pie and bar-of-pie charts, returns or sets the space between the primary and secondary sections of the chart. Read/write  **Long** .


## Syntax

 _expression_ . **GapWidth**

 _expression_ A variable that represents a **[ChartGroup](chartgroup-object-word.md)** object.


## Remarks

The value of this property must be between 0 and 500.


## Example

The following example sets the space between column clusters for the first chart in the active document to be 50 percent of the column width.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.ChartGroups(1).GapWidth = 50 
 End If 
End With
```


## See also


#### Concepts


[ChartGroup Object](chartgroup-object-word.md)

