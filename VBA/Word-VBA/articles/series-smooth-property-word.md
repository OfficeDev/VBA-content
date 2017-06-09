---
title: Series.Smooth Property (Word)
keywords: vbawd10.chm123732131
f1_keywords:
- vbawd10.chm123732131
ms.prod: word
api_name:
- Word.Series.Smooth
ms.assetid: 9360e311-566f-e173-b5e3-ed3790c007fc
ms.date: 06/08/2017
---


# Series.Smooth Property (Word)

 **True** if curve smoothing is enabled for the line chart or scatter chart. Read/write **Boolean** .


## Syntax

 _expression_ . **Smooth**

 _expression_ A variable that represents a **[Series](series-object-word.md)** object.


## Remarks

This property applies only to line and scatter charts. 


## Example

The following example enables curve smoothing for series one for the first chart in the active document. You should run the example on a 2-D line chart.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1).Smooth = True 
 End If 
End With
```


## See also


#### Concepts


[Series Object](series-object-word.md)

