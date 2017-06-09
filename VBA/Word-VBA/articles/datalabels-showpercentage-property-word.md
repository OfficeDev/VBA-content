---
title: DataLabels.ShowPercentage Property (Word)
keywords: vbawd10.chm207489001
f1_keywords:
- vbawd10.chm207489001
ms.prod: word
api_name:
- Word.DataLabels.ShowPercentage
ms.assetid: d13c6988-d751-e084-8fc0-830cc1382906
ms.date: 06/08/2017
---


# DataLabels.ShowPercentage Property (Word)

 **True** to display the percentage value for the data labels on a chart. **False** to hide the value. Read/write **Boolean** .


## Syntax

 _expression_ . **ShowPercentage**

 _expression_ A variable that represents a **[DataLabels](datalabels-object-word.md)** object.


## Example

The following example enables the percentage value to be shown for the data labels of the first series on the first chart.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1).DataLabels. _ 
 ShowPercentage = True 
 End If 
End With
```


## See also


#### Concepts


[DataLabels Object](datalabels-object-word.md)

