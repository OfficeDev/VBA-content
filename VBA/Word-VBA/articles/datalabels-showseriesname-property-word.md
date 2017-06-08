---
title: DataLabels.ShowSeriesName Property (Word)
keywords: vbawd10.chm207488998
f1_keywords:
- vbawd10.chm207488998
ms.prod: word
api_name:
- Word.DataLabels.ShowSeriesName
ms.assetid: 51064a11-512b-d49d-86c1-1839da0576a4
ms.date: 06/08/2017
---


# DataLabels.ShowSeriesName Property (Word)

 **True** to show the series name for the data labels on a chart. **False** to hide the name. Read/write **Boolean** .


## Syntax

 _expression_ . **ShowSeriesName**

 _expression_ A variable that represents a **[DataLabels](datalabels-object-word.md)** object.


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


[DataLabels Object](datalabels-object-word.md)

