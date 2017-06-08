---
title: DataLabels.ShowBubbleSize Property (Word)
keywords: vbawd10.chm207489002
f1_keywords:
- vbawd10.chm207489002
ms.prod: word
api_name:
- Word.DataLabels.ShowBubbleSize
ms.assetid: 3cec847e-ca5f-3062-9049-74d45834f861
ms.date: 06/08/2017
---


# DataLabels.ShowBubbleSize Property (Word)

 **True** to show the bubble size for the data labels on a chart. **False** to hide the bubble size. Read/write **Boolean** .


## Syntax

 _expression_ . **ShowBubbleSize**

 _expression_ A variable that represents a **[DataLabels](datalabels-object-word.md)** object.


## Example

The following example shows the bubble size for the data labels of the first series on the first chart.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1).DataLabels. _ 
 ShowBubbleSize = True 
 End If 
End With
```


## See also


#### Concepts


[DataLabels Object](datalabels-object-word.md)

