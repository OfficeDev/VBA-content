---
title: DataLabels.ShowRange Property (Word)
keywords: vbawd10.chm207489005
f1_keywords:
- vbawd10.chm207489005
ms.prod: word
ms.assetid: 79789465-c1f7-c3ad-7838-b1d795e6b997
ms.date: 06/08/2017
---


# DataLabels.ShowRange Property (Word)

Set to  **True** to display the **Value From Cells** range field in all the chart data labels for a specified chart. Set to **False** to hide that field. Read/write **Boolean**.


## Syntax

 _expression_ . **ShowRange**

 _expression_ A variable that represents a **DataLabels** object.


## Example

The following example displays the  **Value From Cells** range field for all the data labels of the first series in the first chart.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1).DataLabels.ShowRange = True 
 End If 
End With
```


## Property value

 **BOOL**


## See also


#### Concepts


[DataLabels Object](datalabels-object-word.md)

