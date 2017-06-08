---
title: DataLabel.ShowSeriesName Property (Word)
keywords: vbawd10.chm233900006
f1_keywords:
- vbawd10.chm233900006
ms.prod: word
api_name:
- Word.DataLabel.ShowSeriesName
ms.assetid: 6d2a8c88-be7b-711b-1f09-6bf985906fc6
ms.date: 06/08/2017
---


# DataLabel.ShowSeriesName Property (Word)

 **True** to show the series name for the data labels on a chart. **False** to hide the series name. Read/write **Boolean** .


## Syntax

 _expression_ . **ShowSeriesName**

 _expression_ A variable that represents a **[DataLabel](datalabel-object-word.md)** object.


## Example

The following example enables the series name to be shown for the data labels of the first series on the first chart.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1).DataLabels. _ 
 ShowSeriesName = True 
 End If 
End With
```


## See also


#### Concepts


[DataLabel Object](datalabel-object-word.md)

