---
title: DataLabels.ShowLegendKey Property (Word)
keywords: vbawd10.chm207487147
f1_keywords:
- vbawd10.chm207487147
ms.prod: word
api_name:
- Word.DataLabels.ShowLegendKey
ms.assetid: aeacb32a-8ec0-993c-d57c-7df37a164951
ms.date: 06/08/2017
---


# DataLabels.ShowLegendKey Property (Word)

 **True** if the data label legend key is visible. Read/write **Boolean** .


## Syntax

 _expression_ . **ShowLegendKey**

 _expression_ A variable that represents a **[DataLabels](datalabels-object-word.md)** object.


## Example

The following example sets the data labels for series one of the first chart in the active document to show values and the legend key.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1).DataLabels. _ 
 ShowLegendKey = True 
 .Chart.SeriesCollection(1).DataLabels.Type = xlShowValue 
 End If 
End With
```


## See also


#### Concepts


[DataLabels Object](datalabels-object-word.md)

