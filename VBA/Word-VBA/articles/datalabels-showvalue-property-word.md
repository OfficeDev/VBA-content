---
title: DataLabels.ShowValue Property (Word)
keywords: vbawd10.chm207489000
f1_keywords:
- vbawd10.chm207489000
ms.prod: word
api_name:
- Word.DataLabels.ShowValue
ms.assetid: 3c016afc-17b2-78cd-8964-584e8d86d552
ms.date: 06/08/2017
---


# DataLabels.ShowValue Property (Word)

 **True** to display the data label values for a specified chart. **False** to hide the values. Read/write **Boolean** .


## Syntax

 _expression_ . **ShowValue**

 _expression_ A variable that represents a **[DataLabels](datalabels-object-word.md)** object.


## Example

The following example enables the value to be shown for the data labels of the first series in the first chart.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1).DataLabels. _ 
 ShowValue = True 
 End If 
End With
```


## See also


#### Concepts


[DataLabels Object](datalabels-object-word.md)

