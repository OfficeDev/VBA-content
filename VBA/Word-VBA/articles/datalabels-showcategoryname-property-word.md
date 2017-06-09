---
title: DataLabels.ShowCategoryName Property (Word)
keywords: vbawd10.chm207488999
f1_keywords:
- vbawd10.chm207488999
ms.prod: word
api_name:
- Word.DataLabels.ShowCategoryName
ms.assetid: 725deb0e-0b55-8c3d-7893-46d9c25e7b0d
ms.date: 06/08/2017
---


# DataLabels.ShowCategoryName Property (Word)

 **True** to display the category name for the data labels on a chart. **False** to hide the name. Read/write **Boolean** .


## Syntax

 _expression_ . **ShowCategoryName**

 _expression_ A variable that represents a **[DataLabels](datalabels-object-word.md)** object.


## Example

The following example shows the category name for the data labels of the first series on the first chart.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1).DataLabels. _ 
 ShowCategoryName = True 
 End If 
End With
```


## See also


#### Concepts


[DataLabels Object](datalabels-object-word.md)

